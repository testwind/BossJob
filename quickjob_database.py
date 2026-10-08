import glob
import json
import logging
import os
import re
import sqlite3
from datetime import datetime
from typing import Dict, List

from quickjob_config import Config


logger = logging.getLogger('quickjob_apply')


class SQLiteManager:
    """SQLite职位详情存储与历史JSON导入。"""

    META_FIELDS = ['最早采集时间', '最新采集时间', '最早联系时间', '最新联系时间']
    DATETIME_FIELDS = set(META_FIELDS)

    def __init__(self, config: Config, database_file: str = None):
        self.config = config
        self.database_file = self._resolve_database_file(database_file)
        self.connection = sqlite3.connect(self.database_file)
        self.connection.row_factory = sqlite3.Row
        self.create_table()

    def _resolve_database_file(self, database_file: str = None) -> str:
        database_file = database_file or self.config.sqlite_file
        if not os.path.isabs(database_file):
            database_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), database_file)
        os.makedirs(os.path.dirname(database_file), exist_ok=True)
        return database_file

    @property
    def all_fields(self) -> List[str]:
        return self.config.FIELD_ORDER + self.META_FIELDS

    @staticmethod
    def _quote_identifier(identifier: str) -> str:
        return '"' + identifier.replace('"', '""') + '"'

    def create_table(self):
        columns = []
        for field in self.all_fields:
            field_type = 'DATETIME' if field in self.DATETIME_FIELDS else 'TEXT'
            columns.append(f'{self._quote_identifier(field)} {field_type}')
        self.connection.execute(
            f'CREATE TABLE IF NOT EXISTS "job_detail" ({", ".join(columns)})'
        )
        self.connection.execute(
            'CREATE INDEX IF NOT EXISTS "idx_job_detail_job_id" '
            'ON "job_detail" ("职位唯一ID")'
        )
        self.connection.commit()

    @staticmethod
    def _serialize_value(value) -> str:
        if value is None:
            return ''
        if isinstance(value, (list, dict)):
            return json.dumps(value, ensure_ascii=False)
        return str(value)

    @staticmethod
    def _parse_datetime(value):
        if not value:
            return None
        if isinstance(value, datetime):
            return value
        value = str(value).strip()
        for date_format in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y%m%d_%H%M%S', '%Y%m%d'):
            try:
                return datetime.strptime(value, date_format)
            except ValueError:
                continue
        raise ValueError(f'无法解析日期时间: {value}')

    def _record_values(self, record: Dict) -> Dict:
        values = {
            field: self._serialize_value(record.get(field, ''))
            for field in self.config.FIELD_ORDER
        }
        values['数据采集时间'] = self._serialize_value(record.get('数据采集时间', ''))
        return values

    def upsert_record(self, record: Dict) -> str:
        values = self._record_values(record)
        job_id = values.get('职位唯一ID', '')
        collected_at = self._parse_datetime(values.get('数据采集时间')) if values.get('数据采集时间') else None

        existing = None
        if job_id:
            existing = self.connection.execute(
                'SELECT rowid, "最早采集时间", "最新采集时间", '
                '"最早联系时间", "最新联系时间" '
                'FROM "job_detail" WHERE "职位唯一ID" = ?',
                (job_id,)
            ).fetchone()

        if existing:
            earliest = self._parse_datetime(existing['最早采集时间']) if existing['最早采集时间'] else None
            latest = self._parse_datetime(existing['最新采集时间']) if existing['最新采集时间'] else None
            if collected_at and (earliest is None or collected_at < earliest):
                earliest = collected_at
            if collected_at and (latest is None or collected_at > latest):
                latest = collected_at
            values['最早采集时间'] = earliest.strftime('%Y-%m-%d %H:%M:%S') if earliest else ''
            values['最新采集时间'] = latest.strftime('%Y-%m-%d %H:%M:%S') if latest else ''
            values['最早联系时间'] = existing['最早联系时间'] or ''
            values['最新联系时间'] = existing['最新联系时间'] or ''
            assignments = ', '.join(
                f'{self._quote_identifier(field)} = ?'
                for field in self.config.FIELD_ORDER + self.META_FIELDS
            )
            parameters = [values[field] for field in self.config.FIELD_ORDER + self.META_FIELDS]
            self.connection.execute(
                f'UPDATE "job_detail" SET {assignments} WHERE rowid = ?',
                parameters + [existing['rowid']]
            )
            return 'updated'

        values['最早采集时间'] = collected_at.strftime('%Y-%m-%d %H:%M:%S') if collected_at else ''
        values['最新采集时间'] = values['最早采集时间']
        values['最早联系时间'] = ''
        values['最新联系时间'] = ''
        fields = self.all_fields
        placeholders = ', '.join('?' for _ in fields)
        self.connection.execute(
            f'INSERT INTO "job_detail" ({", ".join(self._quote_identifier(field) for field in fields)}) '
            f'VALUES ({placeholders})',
            [values[field] for field in fields]
        )
        return 'inserted'

    def upsert_records(self, records: List[Dict]) -> Dict[str, int]:
        statistics = {'inserted': 0, 'updated': 0, 'failed': 0}
        try:
            with self.connection:
                for record in records:
                    try:
                        result = self.upsert_record(record)
                        statistics[result] += 1
                    except Exception as error:
                        statistics['failed'] += 1
                        logger.warning(f'写入SQLite记录失败: {error}')
        except Exception:
            self.connection.rollback()
            raise
        return statistics

    def _parse_filter_time(self, value):
        return self._parse_datetime(value) if value else None

    def _in_time_range(self, value, start_time, end_time) -> bool:
        return (
            (start_time is None or value >= start_time)
            and (end_time is None or value <= end_time)
        )

    def batch_import_json(self, start_time: str = None, end_time: str = None,
                          job_name: str = None) -> Dict[str, int]:
        """按开始时间、结束时间和职位名称，按文件时间顺序导入最终JSON。"""
        start = self._parse_filter_time(start_time)
        end = self._parse_filter_time(end_time)
        if start and end and start > end:
            raise ValueError('开始时间不能晚于结束时间')

        root = os.path.abspath(self.config.data_root_dir)
        city_pattern = re.compile(rf'^{re.escape(self.config.city_name)}_(\d{{8}})$')
        input_files = []
        if os.path.isdir(root):
            for city_dir in os.listdir(root):
                city_match = city_pattern.match(city_dir)
                if not city_match:
                    continue
                directory_date = datetime.strptime(city_match.group(1), '%Y%m%d')
                if start and directory_date.date() < start.date():
                    continue
                if end and directory_date.date() > end.date():
                    continue
                date_root = os.path.join(root, city_dir)
                for current_job in os.listdir(date_root):
                    if job_name and current_job != job_name:
                        continue
                    job_dir = os.path.join(date_root, current_job)
                    if not os.path.isdir(job_dir):
                        continue
                    file_pattern = f'{self.config.platform}_{self.config.city_name}_{current_job}_*.json'
                    for file_path in glob.glob(os.path.join(job_dir, file_pattern)):
                        match = re.search(r'_(\d{8}_\d{6})\.json$', os.path.basename(file_path))
                        if not match:
                            continue
                        file_time = datetime.strptime(match.group(1), '%Y%m%d_%H%M%S')
                        if self._in_time_range(file_time, start, end):
                            input_files.append((file_time, file_path))

        statistics = {'files': 0, 'records': 0, 'inserted': 0, 'updated': 0, 'failed': 0}
        for _, file_path in sorted(input_files, key=lambda item: (item[0], item[1])):
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    records = json.load(file)
                if not isinstance(records, list):
                    raise ValueError('JSON根节点不是数组')
                result = self.upsert_records(records)
                statistics['files'] += 1
                statistics['records'] += len(records)
                for key in ('inserted', 'updated', 'failed'):
                    statistics[key] += result[key]
                logger.info(
                    f"SQLite导入文件: {os.path.basename(file_path)} | "
                    f"总记录数: {len(records)} | 新插入数量: {result['inserted']} | "
                    f"更新记录数量: {result['updated']} | 失败数量: {result['failed']}"
                )
            except Exception as error:
                statistics['failed'] += 1
                logger.warning(
                    f"SQLite导入文件: {os.path.basename(file_path)} | "
                    f"总记录数: 0 | 新插入数量: 0 | 更新记录数量: 0 | 失败数量: 1 | "
                    f"错误: {error}"
                )
        return statistics

    def close(self):
        self.connection.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
