from dataclasses import dataclass


@dataclass
class Config:
    """统一配置类"""
    platform: str = "快速求职"  # 平台名称，用于实际输出文件名
    platform_url: str = "https://www.zhipin.com"    # 平台网址
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
    cookie_refresh: bool = False  # 是否重新登录{平台}刷新Cookie信息（点开登录界面，扫码登录后就可以不用管了）
    cookie_wait: int = 5  # 登录{平台}页面等待时间
    cookie_file: str = 'cookies.json'  # Cookie文件

    # 目录配置
    data_root_dir: str = 'output'  # 数据根目录
    progress_dir: str = '留存'  # 进度文件存放子目录
    sqlite_file: str = 'quickjob.sqlite3'  # SQLite文件名，默认放在脚本同级目录

    FIELD_ORDER = [
        '序号', '数据列表时间', '数据采集时间', '职位状态', '职位标题', '薪资-0', '薪资', 'tag-icon', '全部标签', '工作城市', '工作区域', '工作地点',
        '经验要求', '学历要求', '岗位标签', '职位描述', '职位详情链接', '职位唯一ID', '公司名称',
        '公司详情链接', '融资情况', '公司规模', '所属行业', '工商-公司名称', '工商-法定代表人', '工商-成立日期',
        '工商-企业类型', '工商-经营状态', '工商-注册资金', '工商-工作地址', '招聘负责人', '活跃状态', '招聘者职位'
    ]
