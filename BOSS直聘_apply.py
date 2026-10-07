"""
Time   : 2026/03/02 14:46 周一
Author : Leo
Version: v2.0
Desc   : Boss直聘数据采集工具 - DrissionPage自动化实现

         【核心功能】
         ⚡ 职位搜索 - 支持城市代码自动获取，智能滚动加载，自动去重
         📊 详情采集 - 职位描述、薪资、经验学历、公司信息、招聘负责人等20+字段
         🔄 断点续传 - 中断后可无缝继续
         💾 数据存储 - JSON + Excel双格式导出
         🔍 HTML调试 - 关键步骤保存页面源码，方便定位元素

         【运行模式】
         1. 完整模式 - 重新搜索职位并采集详情（自动去重）
         2. 续传模式 - 按最新时间戳批次继续采集
         3. 进度转JSON和Excel - 导出最新时间戳批次的进度

         【配置参数】
         - 职位名称、城市、滚动次数、延迟时间、Cookie管理等均在Config类中可调
         - Cookie失效后请登录刷新（cookie_refresh=True）

         【输出结构】
         output/
         └─ {城市}_{日期}/
            └─ {职位}/
               ├─ {职位}_list_{时间戳}.json                    # 职位列表
               ├─ BOSS直聘_{城市}_{职位}_{时间戳}.json          # 详情数据JSON
               ├─ BOSS直聘_{城市}_{职位}_{时间戳}.xlsx          # 详情数据Excel
               ├─ html_debug/                                   # HTML源码调试目录
               │  ├─ 01_首页_{时间戳}.html
               │  ├─ 02_搜索结果页_{时间戳}.html
               │  ├─ 03_滚动加载后_{时间戳}.html
               │  └─ 04_职位详情_{职位ID}_{时间戳}.html
               └─ 留存/
             ├─ {职位}_{时间戳}_progress_batch_001.json    # 第1批新增数据
             └─ {职位}_{时间戳}_progress_batch_002.json    # 第2批新增数据

           说明：同一个职位标签的一次采集使用同一个时间戳，贯穿列表、详情和进度文件。
         每个进度文件只保存当前批次新增成功的数据，不保存final文件。
         所有批次尝试完成后，合并生成详情JSON和Excel；失败或跳过的职位不阻塞完成判断。
"""

import json
import time
import logging
import os
import random
import glob
import re
import requests
from datetime import datetime
from typing import List, Dict, Optional, Set
from dataclasses import dataclass

import pandas as pd
from DrissionPage import ChromiumPage


# ==================== 配置区域（所有可调参数统一在这里修改）====================
@dataclass
class Config:
    """统一配置类"""
    # 搜索配置
    job_name: str = "数据"  # 搜索的职位名称
    city_name: str = "上海"  # 搜索城市名称
    city_code: str = ""  # 城市代码（自动获取）
    search_max_scrolls: int = 30  # 搜索页最大滚动次数（当前平台滚动最多加载300个职位卡片）
    search_wait_time: int = 60  # 搜索页等待时间，此时可以筛选条件（例如：学历，工作经验）

    # 详情页采集配置
    detail_base_delay: float = 2.0  # 详情页基础延迟（秒）
    detail_random_delay: float = 1.0  # 详情页随机延迟范围（±秒）
    detail_save_interval: int = 10  # 每多少个详情页保存一次进度

    # HTML调试配置
    save_html_debug: bool = True  # 是否保存HTML源码用于调试
    max_html_save: int = 5  # 最多保存多少个详情页HTML（避免太多文件）

    # Cookie配置
    cookie_refresh: bool = False  # 是否重新登录BOSS直聘刷新Cookie信息（点开登录界面，扫码登录后就可以不用管了）
    cookie_wait: int = 5  # 登录BOSS直聘页面等待时间
    cookie_file: str = 'cookies.json'  # Cookie文件

    # 目录配置
    data_root_dir: str = 'output'  # 数据根目录
    progress_dir: str = '留存'  # 进度文件存放子目录


# ==================== 工具类 ====================

class LogManager:
    """日志管理器 - 单例模式确保日志只初始化一次"""

    _initialized = False
    _logger = None

    @staticmethod
    def get_logger(name: str = None, log_dir: str = "logs"):
        """设置日志记录器"""
        if LogManager._initialized and LogManager._logger:
            return LogManager._logger

        # 自动获取调用脚本名
        if name is None:
            import inspect
            frame = inspect.stack()[1]
            module = inspect.getmodule(frame[0])
            script_name = os.path.splitext(os.path.basename(module.__file__))[0] if module else "boss_scraper"
        else:
            script_name = name

        logger = logging.getLogger(script_name)
        logger.setLevel(logging.DEBUG)

        if not logger.handlers:
            # 定义日志格式，包含时间戳、文件名、函数名、行号和日志消息
            # formatter = logger.Formatter('[%(asctime)s-%(filename)s][%(funcName)s-%(lineno)d]--%(message)s',datefmt='%Y-%m-%d %H:%M:%S')
            formatter = logging.Formatter(
                '[%(asctime)s][%(funcName)s-%(lineno)d][%(levelname)s] - %(message)s', datefmt='%Y-%m-%d %H:%M:%S'
            )

            # 控制台输出
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)

            # 文件输出（按天切割）
            os.makedirs(log_dir, exist_ok=True)
            from logging.handlers import TimedRotatingFileHandler
            file_handler = TimedRotatingFileHandler(
                filename=os.path.join(log_dir, f"{script_name}.txt"),
                when="midnight",
                interval=1,
                backupCount=7,
                encoding="utf-8",
            )
            file_handler.setFormatter(formatter)
            file_handler.suffix = "%Y%m%d.txt"
            logger.addHandler(file_handler)

            logger.info(f"==================== 日志系统初始化完成 ====================")
            # logger.info(f"主日志文件: {script_name}.txt（按天切片，保留7天）")

        LogManager._initialized = True
        LogManager._logger = logger
        return logger


class DelayManager:
    """延迟管理器 - 智能随机延迟，避免被反爬"""

    def __init__(self, base_delay: float = 2.0, random_range: float = 1.0):
        self.base_delay = base_delay
        self.random_range = random_range

    def wait_before_request(self):
        """请求前等待（随机延迟）"""
        delay = random.uniform(
            max(0.5, self.base_delay - self.random_range),
            self.base_delay + self.random_range
        )
        time.sleep(delay)
        return delay


class FileManager:
    """文件管理器 - 统一处理文件读写、路径管理"""

    def __init__(self, config: Config):
        self.config = config
        self.data_dir = self._create_data_dir()
        self.progress_dir = self._create_progress_dir()
        self.html_debug_dir = self._create_html_debug_dir()

    def _create_data_dir(self) -> str:
        """创建数据目录：output/{城市}_{日期}/{职位}/"""
        today = datetime.now().strftime('%Y%m%d')
        city_date_dir = os.path.join(self.config.data_root_dir, f"{self.config.city_name}_{today}")
        data_dir = os.path.join(city_date_dir, self.config.job_name)
        os.makedirs(data_dir, exist_ok=True)
        return data_dir

    def _create_progress_dir(self) -> str:
        """创建进度文件存放目录：output/{城市}_{日期}/{职位}/留存/"""
        progress_dir = os.path.join(self.data_dir, self.config.progress_dir)
        os.makedirs(progress_dir, exist_ok=True)
        return progress_dir

    def _create_html_debug_dir(self) -> str:
        """创建HTML调试目录：output/{城市}_{日期}/{职位}/html_debug/"""
        if self.config.save_html_debug:
            html_dir = os.path.join(self.data_dir, "html_debug")
            os.makedirs(html_dir, exist_ok=True)
            return html_dir
        else:
            logger.warning("HTML调试模式未开启，跳过创建 html_debug 目录")
            return ''

    def get_progress_file(self, batch_num=None, timestamp: str = None) -> Optional[str]:
        """获取进度文件路径"""
        if batch_num:
            batch_str = f"{batch_num:03d}" if isinstance(batch_num, int) else batch_num
            if not timestamp:
                return None
            return os.path.join(self.progress_dir, f'{self.config.job_name}_{timestamp}_progress_batch_{batch_str}.json')
        else:
            # 查找最新的进度文件
            progress_files = self.get_progress_files(timestamp)
            return max(progress_files, key=os.path.getctime) if progress_files else None

    def get_progress_files(self, timestamp: str = None) -> List[str]:
        """获取指定时间戳的进度文件。"""
        pattern = f'{self.config.job_name}_*_progress_batch_*.json'
        files = glob.glob(os.path.join(self.progress_dir, pattern))
        if timestamp:
            files = [file_path for file_path in files if f'_{timestamp}_progress_batch_' in os.path.basename(file_path)]
        return files

    def get_progress_batch_number(self, file_path: str) -> Optional[int]:
        """从增量进度文件名中获取批次号。"""
        match = re.search(r'_progress_batch_(\d+)\.json$', os.path.basename(file_path))
        return int(match.group(1)) if match else None

    def get_progress_batch_files(self, timestamp: str) -> Dict[int, str]:
        """获取指定时间戳的增量进度文件，按批次号索引。"""
        batch_files = {}
        for file_path in self.get_progress_files(timestamp):
            batch_number = self.get_progress_batch_number(file_path)
            if batch_number is not None:
                batch_files[batch_number] = file_path
        return batch_files

    def find_latest_progress(self, timestamp: str = None) -> Optional[str]:
        """查找最新的进度文件"""
        progress_files = self.get_progress_files(timestamp)
        return max(progress_files, key=os.path.getctime) if progress_files else None

    def find_latest_job_list(self) -> Optional[str]:
        """查找最新的职位列表文件"""
        pattern = os.path.join(self.data_dir, f"{self.config.job_name}_list_????????_??????.json")
        files = glob.glob(pattern)
        return max(files) if files else None

    def find_latest_complete_batch(self):
        """查找同时存在列表和进度文件的最新时间戳批次。"""
        pattern = os.path.join(self.data_dir, f"{self.config.job_name}_list_????????_??????.json")
        list_files = sorted(glob.glob(pattern), reverse=True)
        for list_file in list_files:
            list_name = os.path.basename(list_file)
            timestamp = list_name.rsplit('_list_', 1)[-1].removesuffix('.json')
            progress_file = self.find_latest_progress(timestamp)
            return list_file, progress_file, timestamp
        return None, None, None

    def save_json(self, data: List[Dict], filename: str):
        """保存JSON文件"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_json(self, filename: str) -> List[Dict]:
        """加载JSON文件"""
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"加载JSON文件失败 {filename}: {e}")
            return []

    def save_html(self, html_content: str, step_name: str, job_id: str = None):
        """保存HTML源码到调试目录"""
        if not self.config.save_html_debug:
            return None

        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

            if job_id:
                # 详情页HTML
                filename = os.path.join(self.html_debug_dir, f"{step_name}_{job_id}_{timestamp}.html")
            else:
                # 普通页面HTML
                filename = os.path.join(self.html_debug_dir, f"{step_name}_{timestamp}.html")

            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html_content)

            logger.info(f"HTML源码已保存: {filename}")
            return filename
        except Exception as e:
            logger.warning(f"保存HTML源码失败: {e}")
            return None


class CityCodeManager:
    """城市代码管理器 - 从BOSS直聘API获取城市代码"""

    def __init__(self):
        self.cache = {}  # 缓存已获取的城市代码

    def get_city_code(self, city_name: str) -> Optional[str]:
        """根据城市名称获取城市代码"""
        # 检查缓存
        if city_name in self.cache:
            logger.debug(f"从缓存获取城市 [{city_name}] 代码: {self.cache[city_name]}")
            return self.cache[city_name]

        try:
            logger.debug(f"正在从Boss直聘API获取城市 [{city_name}] 的代码...")
            response = requests.get("https://www.zhipin.com/wapi/zpCommon/data/city.json")

            if response.status_code != 200:
                logger.error(f"请求城市数据失败，状态码: {response.status_code}")
                return None

            data = response.json()

            # 遍历查找城市
            for province in data['zpData']['cityList']:
                for city in province['subLevelModelList']:
                    if city['name'] == city_name:
                        city_code = city['code']
                        self.cache[city_name] = city_code
                        return city_code

                    # 如果城市名带"市"字，尝试去掉再匹配
                    if city_name.endswith('市') and city['name'] == city_name[:-1]:
                        city_code = city['code']
                        self.cache[city_name] = city_code
                        return city_code

            return None
        except Exception as e:
            logger.error(f"获取城市代码失败: {e}")
            return None


# ==================== 页面操作类 ====================
class PageOperator:
    """页面操作器 - 处理浏览器页面操作"""

    def __init__(self, page: ChromiumPage, config: Config, file_manager: FileManager):
        self.page = page
        self.config = config
        self.file_manager = file_manager
        self.delay_manager = DelayManager(
            config.detail_base_delay,
            config.detail_random_delay
        )

    def load_cookie(self) -> bool:
        """加载Cookie"""
        if not os.path.exists(self.config.cookie_file):
            logger.warning(f"Cookie文件不存在: {self.config.cookie_file}")
            return False

        logger.info("开始加载Cookie")
        self.page.get("https://www.zhipin.com")
        time.sleep(2)

        cookies = self.file_manager.load_json(self.config.cookie_file)
        for cookie in cookies:
            self.page.set.cookies(cookie)

        self.page.refresh()
        time.sleep(2)
        logger.info("Cookie加载完成")

        # 保存首页HTML
        if self.config.save_html_debug:
            self.file_manager.save_html(self.page.html, "01_首页_加载Cookie后")

        return True

    def has_login_button(self) -> bool:
        """检查当前页面是否仍显示登录页面或登录按钮"""
        try:
            title = self.page.title or ''
            if '登录' in title:
                return True

            return bool(self.page.ele('css:a[ka="header-login"]', timeout=0))
        except Exception:
            return False

    def get_cookie(self, url: str, timeout=5):
        """获取Cookie（手动登录）"""
        logger.info("请打开登录窗口，扫码登录您的个人账号...")
        self.page.get(url)

        should_login = True

        while should_login:
            logger.info("仍检测到登录按钮，等待登录完成...")
            time.sleep(3)
            should_login = self.has_login_button()
            if not should_login:
                logger.info("检测到没有登录按钮，等待登无变化... %d 秒",timeout)
                time.sleep(timeout)  # 等待确认登录无变化
                should_login = self.has_login_button()

        cookies = self.page.cookies()
        self.file_manager.save_json(cookies, self.config.cookie_file)
        logger.info(f"Cookie已保存到 {self.config.cookie_file}")


# ==================== 数据解析类 ====================
class JobParser:
    """职位解析器 - 解析列表页和详情页数据"""

    # 字段顺序定义（用于统一输出格式）
    FIELD_ORDER = [
        '序号', '数据列表时间', '数据采集时间', '职位状态', '职位标题', '薪资-0', '薪资', 'tag-icon', '全部标签', '工作城市', '工作区域', '工作地点',
        '经验要求', '学历要求', '岗位标签', '职位描述', '职位详情链接', '职位唯一ID', '公司名称',
        '公司详情链接', '融资情况', '公司规模', '所属行业', '工商-公司名称', '工商-法定代表人', '工商-成立日期',
        '工商-企业类型', '工商-经营状态', '工商-注册资金', '工商-工作地址', '招聘负责人', '活跃状态', '招聘者职位'
    ]

    def __init__(self, config: Config):
        self.config = config
        self.job_list_cache = {}  # 缓存列表数据，用于合并详情

    def set_job_list_cache(self, job_list: List[Dict]):
        """设置职位列表缓存，用于合并详情数据"""
        for job in job_list:
            job_id = job.get('职位唯一ID') or self.extract_job_id(job.get('职位详情链接', ''))
            if job_id:
                self.job_list_cache[job_id] = job
        logger.info(f"已缓存 {len(self.job_list_cache)} 个职位列表信息")

    def clean_text(self, text: str) -> str:
        """清洗文本（去除多余空格，保留换行符，并移除BOSS直聘相关品牌字符）"""
        if not text:
            return ''

        # 先按行分割，处理每一行，然后重新组合
        lines = text.split('\n')
        cleaned_lines = []
        for line in lines:
            # 对每一行进行基本的空白字符清理（保留行内的空格，但去除多余空格）
            cleaned_line = ' '.join(line.split())
            if cleaned_line:  # 只处理非空行
                cleaned_lines.append(cleaned_line)
        # 用换行符重新连接
        cleaned = '\n'.join(cleaned_lines)
        # 使用正则表达式匹配 "boss"（不区分大小写），前面可带可不带"来自"，后面可带可不带"直聘"，或者单独匹配 "直聘"
        pattern = re.compile(
            r'(?:来自\s*)?boss\s*(?:直聘)?|直聘',
            re.IGNORECASE  # 忽略大小写
        )  # 匹配各种形式的BOSS直聘及相关组合
        # 替换所有匹配（跨行处理）
        cleaned = pattern.sub('', cleaned)
        # 清理替换后可能留下的多余标点符号（每行单独处理，避免跨行影响）
        lines = cleaned.split('\n')

        final_lines = []
        for line in lines:
            # 清理行内多余的分隔符
            line = re.sub(r'[·\-_]{2,}', '', line)  # 连续多个分隔符
            line = re.sub(r'^[·\-_\s]+|[·\-_\s]+$', '', line)  # 开头或结尾的分隔符
            # 如果行变成了空字符串或只有标点符号，跳过该行
            if line and not all(c in '·-_.,，。、\s' for c in line):
                final_lines.append(line)
        # 用换行符重新连接
        cleaned = '\n'.join(final_lines)

        return cleaned

    def extract_job_id(self, url: str) -> str:
        """从URL中提取职位唯一ID"""
        if not url:
            return ''

        # 提取securityId
        security_match = re.search(r'[?&]securityId=([^&]+)', url)
        if security_match:
            return security_match.group(1)

        # 提取encId (job_detail后的ID)
        enc_match = re.search(r'/job_detail/([^.?]+)', url)
        if enc_match:
            return enc_match.group(1)

        return ''

    def decode_salary(self, salary: str) -> str:
        """解码列表页薪资中的私有区数字字符。"""
        if not salary:
            return ''

        return ''.join(
            str(ord(char) - 0xE031) if 0xE031 <= ord(char) <= 0xE03A else char
            for char in salary
        ).strip()

    def parse_job_list(self, page: ChromiumPage) -> List[Dict]:
        """解析职位列表页（带去重功能）"""
        logger.info("开始解析职位列表")
        # 匹配class包含"rec-job-list"的ul元素，// 表示从根节点开始搜索（从整个文档搜索）
        job_list = page.ele('xpath://ul[contains(@class, "rec-job-list")]', timeout=10)
        if not job_list:
            logger.warning("未找到职位列表")
            return []

        # 获取所有符合条件的职位卡片：在job_list内部搜索所有class包含"job-card-box"的li元素，eles() 返回多个元素（列表），用于定位多个匹配项，.// 表示从当前节点开始搜索
        job_cards = job_list.eles('xpath:.//li[contains(@class, "job-card-box")]', timeout=10)
        logger.info(f"找到 {len(job_cards)} 个职位卡片")

        jobs = []
        seen_ids = set()  # 用于去重的ID集合

        for i, card in enumerate(job_cards, 1):
            try:
                job = self._parse_job_card(card, i)
                if job and job.get('职位详情链接'):
                    job_id = self.extract_job_id(job['职位详情链接'])
                    if job_id:
                        if job_id not in seen_ids:
                            job['职位唯一ID'] = job_id
                            seen_ids.add(job_id)
                            jobs.append(job)
                        else:
                            logger.debug(f"跳过重复职位: {job['职位名称']} (ID: {job_id})")
                    else:
                        # 如果没有提取到ID，用URL去重
                        if job['职位详情链接'] not in seen_ids:
                            seen_ids.add(job['职位详情链接'])
                            jobs.append(job)
            except Exception as e:
                logger.warning(f"解析第{i}个职位卡片失败: {e}")

        logger.info(f"解析完成: 原始 {len(job_cards)} 个，去重后 {len(jobs)} 个")
        return jobs

    def _parse_job_card(self, card, index: int) -> Optional[Dict]:
        """解析单个职位卡片"""
        job = {
            '序号': index,
            '职位名称': '',
            '数据列表时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            '薪资-0': '',
            'tag-icon': '',
            '全部标签': [],
            '公司': '',
            '工作地点': '',
            '经验要求': '',
            '学历要求': '',
            '职位详情链接': '',
            '公司详情链接': '',
            '职位唯一ID': ''
        }

        # 职位名称和链接
        name_elem = self._get_ele(card, 'xpath:.//a[contains(@class, "job-name")]', timeout=0)
        if name_elem:
            job['职位名称'] = self.clean_text(name_elem.text)
            href = name_elem.attr('href')
            if href:
                job['职位详情链接'] = 'https://www.zhipin.com' + href if not href.startswith('http') else href

        # 薪资中的数字使用私有区字符编码，需要先还原为普通数字。
        salary_elem = self._get_ele(card, 'xpath:.//span[contains(@class, "job-salary")]', timeout=0)
        if salary_elem:
            job['薪资-0'] = self.decode_salary(salary_elem.text)

        tag_icon_elem = self._get_ele(card, 'css:img.job-tag-icon', timeout=0)
        if tag_icon_elem:
            job['tag-icon'] = self.clean_text(tag_icon_elem.attr('alt') or '')

        # 公司
        company_elem = self._get_ele(card, 'xpath:.//span[contains(@class, "boss-name")]', timeout=0)
        if company_elem:
            job['公司'] = self.clean_text(company_elem.text)

        # 地点
        loc_elem = self._get_ele(card, 'xpath:.//span[contains(@class, "company-location")]', timeout=0)
        if loc_elem:
            job['工作地点'] = self.clean_text(loc_elem.text)

        # 经验学历标签
        tag_list = self._get_ele(card, 'xpath:.//ul[contains(@class, "tag-list")]', timeout=0)
        if tag_list:
            tags = [self.clean_text(tag.text) for tag in tag_list.eles('tag:li')]
            job['全部标签'] = tags
            if len(tags) > 0:
                job['经验要求'] = tags[0]
            if len(tags) > 1:
                job['学历要求'] = tags[1]

        # 公司链接
        company_link_elem = self._get_ele(card, 'xpath:.//a[contains(@class, "boss-info")]', timeout=0)
        if company_link_elem:
            href = company_link_elem.attr('href')
            if href:
                job['公司详情链接'] = 'https://www.zhipin.com' + href if not href.startswith('http') and not href.startswith('javascript') else href

        return job if job['职位名称'] else None

    def parse_job_detail(self, page: ChromiumPage, index: int, job_url: str = '') -> Dict:
        """解析职位详情页，并与列表数据合并"""
        job_id = self.extract_job_id(job_url)

        # 初始化基础数据
        job = {
            '序号': index,
            '数据列表时间': '',
            '数据采集时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            '职位状态': '',
            '职位标题': '',
            '薪资-0': '',
            '薪资': '',
            'tag-icon': '',
            '全部标签': [],
            '工作城市': '',
            '工作区域': '',
            '工作地点': '',
            '经验要求': '',
            '学历要求': '',
            '岗位标签': '',
            '职位描述': '',
            '职位详情链接': job_url,
            '职位唯一ID': job_id,
            '公司名称': '',
            '融资情况': '',
            '公司规模': '',
            '所属行业': '',
            '公司详情链接': '',
            '工商-公司名称': '',
            '工商-法定代表人': '',
            '工商-成立日期': '',
            '工商-企业类型': '',
            '工商-经营状态': '',
            '工商-注册资金': '',
            '工商-工作地址': '',
            '招聘负责人': '',
            '活跃状态': '',
            '招聘者职位': ''
        }

        # 从缓存中合并列表数据
        if job_id and job_id in self.job_list_cache:
            list_data = self.job_list_cache[job_id]
            job['职位标题'] = list_data.get('职位名称', '')
            job['数据列表时间'] = list_data.get('数据列表时间', '')
            job['薪资-0'] = list_data.get('薪资-0', '')
            job['tag-icon'] = list_data.get('tag-icon', '')
            job['全部标签'] = list_data.get('全部标签', [])
            job['工作区域'] = list_data.get('工作地点', '').split('·')[1] if '·' in list_data.get('工作地点', '') else ''
            job['工作地点'] = list_data.get('工作地点', '') if '·' in list_data.get('工作地点', '') else ''
            job['公司名称'] = list_data.get('公司', '')
            job['公司详情链接'] = list_data.get('公司详情链接', '')
            if list_data.get('经验要求'):
                job['经验要求'] = list_data.get('经验要求', '')
            if list_data.get('学历要求'):
                job['学历要求'] = list_data.get('学历要求', '')

        try:
            # 基本信息区域
            info_primary = page.ele('xpath://div[@class="info-primary"]', timeout=5)
            if info_primary:
                job['职位状态'] = self._get_text(info_primary, 'xpath:.//div[contains(@class, "job-status")]')
                job['职位标题'] = self._get_text(info_primary, 'xpath:.//h1[@title]')
                # 薪资（列表里的薪资是无效的）
                job['薪资'] = self._get_text(info_primary, 'xpath:.//span[contains(@class, "salary")]')

                # 城市、经验、学历
                info_text = info_primary.ele('tag:p')
                if info_text:
                    job['工作城市'] = self._get_text(info_text, 'xpath:.//a[contains(@class, "text-city")]')

                    exp = self._get_ele(info_text,'xpath:.//span[contains(@class, "text-experiece")]')
                    if exp:
                        job['经验要求'] = self.clean_text(exp.text)

                    edu = self._get_ele(info_text,'xpath:.//span[contains(@class, "text-degree")]')
                    if edu:
                        job['学历要求'] = self.clean_text(edu.text)

            # 职位描述和标签
            job_detail = page.ele('xpath://div[@class="job-detail-section"]', timeout=0)
            if job_detail:
                job['职位描述'] = self._get_text(job_detail, 'xpath:.//div[contains(@class, "job-sec-text")]')
                # 技能标签
                skills = []
                keyword_list = self._get_ele( job_detail, 'xpath:.//ul[contains(@class, "job-keyword-list")]', timeout=0)
                if keyword_list:
                    for item in keyword_list.eles('xpath:.//li', timeout=0):
                        skills.append(self.clean_text(item.text))
                if skills:
                    job['岗位标签'] = '、'.join(skills)

            # 公司信息
            company = page.ele('xpath://div[@class="sider-company"]', timeout=0)
            if company:
                job['公司名称'] = self._get_text(company, 'xpath:.//a[@ka="job-detail-company_custompage"]')
                job['所属行业'] = self._get_text(company, 'xpath:.//a[@ka="job-detail-brandindustry"]')
                # 在company中查找内部包含class为"icon-stage"的i标签的p元素
                job['融资情况'] = self._get_text(company, 'xpath:.//p[i[contains(@class, "icon-stage")]]')
                job['公司规模'] = self._get_text(company, 'xpath:.//p[i[contains(@class, "icon-scale")]]')

            # 工商信息
            business_info = self._get_ele(page,
                'xpath://div[contains(@class, "business-info-box")]', timeout=0
            )
            business_fields = {
                'company-name': '工商-公司名称',
                'company-user': '工商-法定代表人',
                'res-time': '工商-成立日期',
                'company-type': '工商-企业类型',
                'manage-state': '工商-经营状态',
                'company-fund': '工商-注册资金',
            }
            if business_info:
                for class_name, field_name in business_fields.items():
                    item = business_info.ele(
                        f'css:li.{class_name}', timeout=0
                    )
                    if item:
                        label = item.ele('tag:span', timeout=0)
                        value = item.text
                        if label:
                            value = value.replace(label.text, '', 1)
                        job[field_name] = self.clean_text(value)

            # 工商信息下方的公司工作地址
            address = self._get_ele(page,
                'css:div.company-address div.location-address', timeout=0
            )
            if address:
                job['工商-工作地址'] = self.clean_text(address.text)

            # 岗位标签
            welfare = []
            job_tags = self._get_ele(page,'xpath://div[contains(@class, "job-tags")]', timeout=0)
            if job_tags:
                for span in job_tags.eles('xpath:.//span', timeout=0):
                    if span.text.strip():
                        welfare.append(self.clean_text(span.text))
            job['岗位标签'] = '、'.join(welfare)

            # 招聘负责人信息（姓名、活跃状态、招聘者职位）
            boss = page.ele('xpath://div[@class="job-boss-info"]', timeout=0)
            if boss:
                # 获取姓名元素
                name_elem = self._get_text(boss, 'xpath:.//h2[@class="name"]')
                if name_elem:
                    full_text = name_elem.strip()
                    # 获取状态元素（可能是boss-active-time或boss-online-tag）
                    status_elem = self._get_text(boss, 'xpath:.//span[contains(@class, "boss-active-time") or contains(@class, "boss-online-tag")]')
                    status = status_elem.strip() if status_elem else ''
                    # 如果状态在姓名里，从姓名中移除
                    name = full_text.replace(status, '').strip() if status and status in full_text else full_text

                    job['招聘负责人'] = name
                    job['活跃状态'] = status

                boss_attr = self._get_text(boss, 'xpath:.//div[@class="boss-info-attr"]')
                if boss_attr:
                    job['招聘者职位'] = boss_attr.split('·', 1)[-1].strip() if '·' in boss_attr else boss_attr

        except Exception as e:
            logger.warning(f"解析职位详情失败: {e}")

        return job

    def _get_ele(self, element, selector: str, timeout: float = 0):
        """安全获取元素文本"""
        try:
            if not element:
                return None
            elem = element.ele(selector, timeout=timeout)
            if not elem:
                return None

            kanzhun_found = False
            for span in elem.eles('tag:span', timeout=timeout):
                if span.text.strip() == 'kanzhun':
                    span.run_js('this.remove()')
                    kanzhun_found = True
            if kanzhun_found:
                logger.debug(f'清洗：kanzhun')
            return elem
        except:
            return None


    def _get_eles(self, element, selector: str, timeout: float = 0):
        """安全获取元素数组"""
        try:
            if not element:
                return []
            elems = element.eles(selector, timeout=timeout)
            if not elems:
                return []

            kanzhun_found = False
            for elem in elems:
                for span in elem.eles('tag:span', timeout=timeout):
                    if span.text.strip() == 'kanzhun':
                        span.run_js('this.remove()')
                        kanzhun_found = True
            if kanzhun_found:
                logger.debug(f'清洗：kanzhun')
            return elems
        except:
            return []

    def _get_text(self, element, selector: str) -> str:
        """安全获取元素文本"""
        try:
            elem = self._get_ele(element, selector)
            if not elem:
                return ''

            return elem.text
        except:
            return ''


# ==================== 采集器 ====================
class JobScraper:
    """职位采集器 - 核心采集逻辑"""

    def __init__(self, page: ChromiumPage, config: Config):
        self.page = page  # 页面对象
        self.config = config  # 配置对象
        self.parser = JobParser(config)  # 解析器对象
        self.file_manager = FileManager(config)  # 文件管理器对象
        self.page_operator = PageOperator(page, config, self.file_manager)  # 页面操作器对象
        self.delay_manager = DelayManager(
            config.detail_base_delay,
            config.detail_random_delay
        )  # 延迟管理器对象

        # 实例变量，存储职位列表和链接
        self.job_list: List[Dict] = []  # 职位列表数据
        self.job_links: List[str] = []  # 职位链接列表
        self.current_job_details: List[Dict] = []  # 当前已采集的详情数据
        self.html_save_count = 0  # 已保存的详情页HTML数量

    def save_page_html(self, step_name: str, job_id: str = None):
        """保存当前页面HTML源码"""
        if not self.config.save_html_debug:
            return

        # 限制保存的详情页HTML数量
        if job_id and self.html_save_count >= self.config.max_html_save:
            return

        filename = self.file_manager.save_html(self.page.html, step_name, job_id)
        if filename and job_id:
            self.html_save_count += 1

    def navigate_to_jobs_page(self):
        """第一步：导航到职位页"""
        logger.info(f"=== 步骤1: 正在导航到职位页... [{self.config.job_name}] 城市: [{self.config.city_name}] ===")

        try:
            # 访问职位页
            logger.info("正在访问BOSS直聘职位页...")
            self.page.get("https://www.zhipin.com/web/geek/jobs")
            # 等待搜索框出现（匹配class或placeholder符合的input元素）
            self.page.wait.ele_displayed('xpath://div[@class="expect-list has-add no-part"]', timeout=10)

        except Exception as e:
            logger.error(f"搜索职位失败: {e}")
            return []

    def get_expected_job_tabs(self):
        """获取期待职位标签"""
        # 定位容器并返回所有标签
        container = self.page.ele('xpath://div[@class="expect-list has-add no-part"]', timeout=5)
        if not container:
            logger.warning("未找到期待职位列表容器")
            return []
        return container.eles('xpath:.//a[contains(@class, "expect-item")]', timeout=5)

    def _scroll_load_for_tab(self):
        """滚动加载更多职位"""
        logger.info(f"开始滚动加载，最大滚动次数: {self.config.search_max_scrolls}")

        last_count = 0
        no_increase = 0

        for i in range(self.config.search_max_scrolls):
            self.page.scroll.to_bottom()
            time.sleep(1.5)  # 等待加载

            # 定位职位列表容器：匹配class包含"rec-job-list"的ul元素
            job_list = self.page.ele('xpath://ul[contains(@class, "rec-job-list")]', timeout=5)
            if job_list:
                # 获取当前页职位数量：在job_list内部搜索所有class包含"job-card-box"的li元素，eles() 返回多个元素（列表）
                cards = job_list.eles('xpath:.//li[contains(@class, "job-card-box")]')
                current = len(cards)
                logger.info(f"第{i + 1}次滚动后: {current}个职位")

                if current > last_count:
                    no_increase = 0
                    last_count = current
                else:
                    no_increase += 1
                    if no_increase >= 3:
                        logger.info("连续3次无新增，停止滚动")
                        break

            time.sleep(1)

    def scrape_expected_jobs_mode(self):
        """【新模式1】遍历期待职位标签进行采集"""
        self.navigate_to_jobs_page()
        initial_tabs = self.get_expected_job_tabs()
        tab_names = []
        for tab in initial_tabs:
            try:
                tab_name = tab.ele('xpath:.//span').text.strip()
            except Exception as e:
                logger.warning(f"读取职位标签名称失败，跳过该标签: {e}")
                continue
            if tab_name:
                tab_names.append(tab_name)

        logger.info(f"找到 {len(tab_names)} 个期待职位标签")

        for i, tab_name in enumerate(tab_names):
            # 详情采集期间页面会跳转，旧标签元素会失效；每轮重新获取当前页面元素。
            if i > 0:
                self.navigate_to_jobs_page()

            tabs = self.get_expected_job_tabs()
            tab = None
            for current_tab in tabs:
                try:
                    current_name = current_tab.ele('xpath:.//span').text.strip()
                except Exception as e:
                    logger.warning(f"读取当前职位标签失败: {e}")
                    continue
                if current_name == tab_name:
                    tab = current_tab
                    break

            if not tab:
                logger.warning(f"未找到职位标签: {tab_name}")
                continue

            logger.info(f"正在采集标签: {tab_name} ({i + 1}/{len(tab_names)})")
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # 更新配置中的职位名称以匹配当前标签
            self.config.job_name = tab_name
            # 重新初始化文件管理器以创建新目录
            self.file_manager = FileManager(self.config)
            
            tab.click()
            time.sleep(3) # 等待页面刷新

            # 当前职位标签已切换完成，此时保存到当前职位对应的目录。
            self.save_page_html("01_职位首页")
            
            self._scroll_load_for_tab()
            
            jobs = self.parser.parse_job_list(self.page)
            if jobs:
                self.job_list = jobs
                self.job_links = [job.get('职位详情链接') for job in jobs if job.get('职位详情链接')]
                
                # 保存列表
                filename = os.path.join(self.file_manager.data_dir, f"{tab_name}_list_{timestamp}.json")
                self.file_manager.save_json(jobs, filename)
                
                # 采集详情
                self.scrape_details(timestamp)
                self.save_results(timestamp)
            
            # 恢复采集器状态
            self.current_job_details = []
            self.job_list = []
            self.job_links = []

    def load_job_list_from_file(self, file_path: str = None) -> bool:
        """从文件加载职位列表"""
        if file_path:
            if not os.path.exists(file_path):
                logger.error(f"文件不存在: {file_path}")
                return False
            jobs = self.file_manager.load_json(file_path)
        else:
            latest_list = self.file_manager.find_latest_job_list()
            if not latest_list:
                logger.error("未找到职位列表文件")
                return False
            logger.info(f"使用最新列表文件: {latest_list}")
            jobs = self.file_manager.load_json(latest_list)

        if not jobs:
            logger.error("职位列表为空")
            return False

        self.job_list = jobs
        self.job_links = [job.get('职位详情链接') for job in jobs if job.get('职位详情链接')]
        logger.info(f"从文件加载了 {len(self.job_links)} 个职位链接")
        return True

    def merge_progress_files(self, timestamp: str) -> List[Dict]:
        """合并指定时间戳下的增量进度文件。"""
        merged_data = {}
        batch_files = self.file_manager.get_progress_batch_files(timestamp)
        for batch_number in sorted(batch_files):
            progress_data = self.file_manager.load_json(batch_files[batch_number])
            for item in progress_data:
                sequence = item.get('序号')
                if sequence:
                    merged_data[sequence] = item
        return [merged_data[sequence] for sequence in sorted(merged_data)]

    def expected_progress_batch_count(self) -> int:
        """根据职位列表数量计算应有的进度批次数。"""
        if not self.job_links:
            return 0
        return (len(self.job_links) + self.config.detail_save_interval - 1) // self.config.detail_save_interval

    def is_progress_complete(self, timestamp: str) -> bool:
        """检查所有职位批次是否均已尝试完成。"""
        batch_files = self.file_manager.get_progress_batch_files(timestamp)
        expected_count = self.expected_progress_batch_count()
        return expected_count > 0 and all(
            batch_number in batch_files for batch_number in range(1, expected_count + 1)
        )

    def progress_processed_indices(self, timestamp: str) -> Set[int]:
        """根据已保存的批次范围获取已经尝试过的职位序号。"""
        processed_indices = set()
        for batch_number in self.file_manager.get_progress_batch_files(timestamp):
            start = (batch_number - 1) * self.config.detail_save_interval + 1
            end = min(batch_number * self.config.detail_save_interval, len(self.job_links))
            processed_indices.update(range(start, end + 1))
        return processed_indices

    def scrape_details(self, timestamp: str) -> List[Dict]:
        """第二步：采集职位详情"""
        if not self.job_links:
            logger.error("没有职位链接可采集")
            return []

        logger.info(f"=== 步骤2: 采集职位详情 ({len(self.job_links)}个) ===")

        # 设置职位列表缓存
        if self.job_list:
            self.parser.set_job_list_cache(self.job_list)

        self.current_job_details = []
        processed_indices: Set[int] = set()
        self.html_save_count = 0  # 重置HTML保存计数

        return self._continue_scraping(processed_indices, timestamp)

    def _continue_scraping(self, processed_indices: Set[int], timestamp: str) -> List[Dict]:
        """继续采集未完成的职位"""
        success_count = len(self.current_job_details)
        fail_count = 0
        batch_records = []

        for i, link in enumerate(self.job_links, 1):
            current_idx = i

            if current_idx in processed_indices:
                # logger.debug(f"跳过第{current_idx}个（已抓取）")
                continue

            logger.info(f"采集 [{current_idx}/{len(self.job_links)}]: {link}")

            # 请求前延迟
            delay = self.delay_manager.wait_before_request()
            logger.debug(f"延迟 {delay:.2f}s")

            try:
                self.page.get(link)  # 请求职位详情页
                time.sleep(1)  # 页面渲染等待

                # 提取职位ID用于文件名
                job_id = self.parser.extract_job_id(link)

                # 保存详情页HTML（前几个用于调试）
                if job_id:
                    self.save_page_html("04_职位详情", job_id)

                job_data = self.parser.parse_job_detail(self.page, current_idx, link)
                self.current_job_details.append(job_data)
                batch_records.append(job_data)
                success_count += 1
                logger.info(f"成功: {job_data.get('职位标题')} | {job_data.get('薪资')} | {job_data.get('经验要求')} | "
                            f"{job_data.get('学历要求')} | {job_data.get('招聘负责人')} | {job_data.get('活跃状态')}")

            except Exception as e:
                logger.error(f"采集失败: {e}")
                fail_count += 1

            # 按间隔保存进度
            if current_idx % self.config.detail_save_interval == 0:
                batch = current_idx // self.config.detail_save_interval
                if batch not in self.file_manager.get_progress_batch_files(timestamp):
                    self._save_progress(batch, timestamp, batch_records)
                batch_records = []

        # 保存最后一个不足整批的批次；空数组也表示该批次已经尝试完成。
        if len(self.job_links) % self.config.detail_save_interval:
            batch = (len(self.job_links) // self.config.detail_save_interval) + 1
            if batch not in self.file_manager.get_progress_batch_files(timestamp):
                self._save_progress(batch, timestamp, batch_records)

        logger.info(f"详情采集完成: 成功{success_count} 失败{fail_count}")
        return self.current_job_details

    def _save_progress(self, batch_num, timestamp: str, records: List[Dict]):
        """保存进度到留存目录"""
        try:
            # 按字段顺序整理
            ordered = []
            for item in records:
                ordered_item = {}
                for field in JobParser.FIELD_ORDER:
                    ordered_item[field] = item.get(field, '')
                ordered.append(ordered_item)

            batch_str = f"{batch_num:03d}" if isinstance(batch_num, int) else batch_num
            filename = os.path.join(
                self.file_manager.progress_dir,
                f"{self.config.job_name}_{timestamp}_progress_batch_{batch_str}.json"
            )
            self.file_manager.save_json(ordered, filename)
            logger.info(f"进度已保存: {filename}")

        except Exception as e:
            logger.error(f"保存进度失败: {e}")

    def save_results(self, timestamp: str):
        """保存最终结果：BOSS直聘_{城市}_{职位}_{时间戳}.json 和 .xlsx"""
        # 最终导出始终以指定时间戳的全部增量进度文件为准。
        self.current_job_details = self.merge_progress_files(timestamp)
        if not self.current_job_details:
            logger.warning("没有成功采集的数据，将保存空的合并结果")

        # 按字段顺序整理
        ordered = []
        for item in self.current_job_details:
            ordered_item = {}
            for field in JobParser.FIELD_ORDER:
                ordered_item[field] = item.get(field, '')
            ordered.append(ordered_item)

        # 保存JSON
        json_file = os.path.join(
            self.file_manager.data_dir,
            f"BOSS直聘_{self.config.city_name}_{self.config.job_name}_{timestamp}.json"
        )
        self.file_manager.save_json(ordered, json_file)

        # 保存Excel
        excel_file = os.path.join(
            self.file_manager.data_dir,
            f"BOSS直聘_{self.config.city_name}_{self.config.job_name}_{timestamp}.xlsx"
        )
        df = pd.DataFrame(ordered)
        df.to_excel(excel_file, index=False)

        logger.info(f"结果已保存: \nJSON: {json_file}\nExcel: {excel_file}")

        # 输出HTML调试文件位置
        if self.config.save_html_debug:
            logger.info(f"HTML调试文件保存在: {self.file_manager.html_debug_dir}")

        return json_file, excel_file


# ==================== 主程序 ====================
def setup_page() -> ChromiumPage:
    """初始化浏览器页面"""
    page = ChromiumPage()  # 创建ChromiumPage对象
    page.set.timeouts(base=30)  # 设置超时时间为30秒
    page.set.window.max()  # 让浏览器窗口占满屏幕，模拟真人操作
    page.set.user_agent(
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36'
    )
    return page


def print_menu():
    """打印菜单"""
    print(f"\n{'=' * 50}")
    print("请选择运行模式：")
    print("1. 完整模式：重新搜索并采集（自动去重）")
    print("2. 续传模式：直接从上次中断处继续（自动检测进度）")
    print("3. 进度转JSON和Excel：导出最新时间戳批次")
    print("0. 退出程序")
    print(f"{'=' * 50}")


def main():
    """主函数"""
    # 初始化配置
    config = Config()

    # 设置日志
    global logger
    logger = LogManager.get_logger()

    # 添加启动日志
    logger.info("=" * 50)
    logger.info(f"Boss直聘爬虫启动 v2.0")
    logger.info(f"配置信息: 职位={config.job_name}, 城市={config.city_name}")
    logger.info(f"HTML调试模式: {'开启' if config.save_html_debug else '关闭'}")

    # 获取城市代码
    if config.city_name and not config.city_code:
        logger.info(f"正在获取城市 [{config.city_name}] 的代码...")
        city_manager = CityCodeManager()
        code = city_manager.get_city_code(config.city_name)
        if code:
            config.city_code = code
            logger.info(f"✓ 城市 [{config.city_name}] 的代码获取成功: {config.city_code}")
        else:
            logger.warning(f"✗ 未找到城市 [{config.city_name}] 的代码，将使用默认搜索")

    logger.info("=" * 50)

    # 初始化浏览器和采集器
    page = setup_page()
    scraper = JobScraper(page, config)
    logger.info(f"数据目录: {scraper.file_manager.data_dir}")
    logger.info(f"进度目录: {scraper.file_manager.progress_dir}")
    if config.save_html_debug:
        logger.info(f"HTML调试目录: {scraper.file_manager.html_debug_dir}")

    try:
        # 是否重新登录BOSS直聘刷新Cookie
        config.cookie_refresh = True
        if config.cookie_refresh:
            scraper.page_operator.get_cookie(url='https://www.zhipin.com', timeout=config.cookie_wait)

        # 加载Cookie
        if not scraper.page_operator.load_cookie():
            logger.error("Cookie加载失败，程序退出")
            return

        while True:
            # 显示菜单
            print_menu()
            mode = input("请输入选择: ").strip()

            if mode == "0":
                logger.info("程序退出")
                break

            if mode == "1":
                # 完整模式：遍历期待职位并采集
                scraper.scrape_expected_jobs_mode()
                logger.info("所有期待职位采集完成")
                return

            elif mode == "2":
                # 使用同一时间戳配对列表文件和进度文件
                latest_list, latest_progress, timestamp = scraper.file_manager.find_latest_complete_batch()
                if not latest_list or not timestamp:
                    logger.error("未找到时间戳列表批次，无法续传，请先运行完整模式生成职位列表")
                    continue

                # 加载职位列表文件
                if not scraper.load_job_list_from_file(latest_list):
                    logger.error("加载职位列表失败")
                    continue

                # 设置职位列表缓存
                if scraper.job_list:
                    scraper.parser.set_job_list_cache(scraper.job_list)
                # 有进度文件，尝试续传
                logger.info(f"检测到时间戳批次 [{timestamp}]，进度文件: {latest_progress or '暂无，开始新采集'}")
                progress_data = scraper.merge_progress_files(timestamp)
                processed_indices = scraper.progress_processed_indices(timestamp)
                scraper.current_job_details = progress_data

                if scraper.is_progress_complete(timestamp):
                    logger.info("所有职位批次均已尝试完成，重新整理结果文件")
                    scraper.save_results(timestamp)
                    return

                logger.info(
                    f"从进度恢复: 已保存 {len(progress_data)} 个成功职位，将继续采集未完成批次")

                # 开始采集
                scraper._continue_scraping(processed_indices, timestamp)
                if scraper.is_progress_complete(timestamp):
                    scraper.save_results(timestamp)
                else:
                    logger.warning("进度批次尚未全部完成，请稍后继续运行续传模式")
                return

            elif mode == "3":
                logger.info("=== 合并最新时间戳批次进度为JSON和Excel ===")
                latest_list, _, timestamp = scraper.file_manager.find_latest_complete_batch()

                if not latest_list or not timestamp:
                    logger.error("未找到时间戳列表批次，请先运行完整模式生成职位列表")
                    continue
                if not scraper.load_job_list_from_file(latest_list):
                    logger.error("加载时间戳列表失败，无法导出结果")
                    continue
                if not scraper.is_progress_complete(timestamp):
                    logger.error("最新时间戳批次尚未完成全部职位尝试，无法导出结果")
                    continue
                logger.info(f"正在合并时间戳 [{timestamp}] 的增量进度文件")
                scraper.current_job_details = scraper.merge_progress_files(timestamp)
                if not scraper.current_job_details:
                    logger.warning("所有职位均采集失败，生成空结果文件")
                if not scraper.current_job_details and scraper.expected_progress_batch_count() == 0:
                    logger.error("职位列表为空，无法导出结果")
                    continue
                scraper.save_results(timestamp)

            else:
                print("无效的选择，请重新输入")
                continue

    except KeyboardInterrupt:
        logger.info("用户中断程序")
    except Exception as e:
        logger.error(f"程序错误: {e}", exc_info=True)
    finally:
        try:
            page.quit()
            logger.info("浏览器已关闭")
        except:
            pass


if __name__ == '__main__':
    main()
