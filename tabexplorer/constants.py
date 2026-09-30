"""全局常量与版本号。"""

# 版本号唯一来源是根目录 TabEx.py（打包脚本与旧版检查更新都读取这一行）
from TabEx import APP_VERSION  # noqa: F401


MAX_CLOSED_TABS_HISTORY = 20  # 关闭标签页历史最大数量（从10增加到20）
MAX_SEARCH_HISTORY = 30  # 搜索历史最大数量（从20增加到30）
MAX_NAVIGATION_HISTORY = 50  # 导航历史最大数量
MAX_ACTIVE_TOASTS = 5  # 同时显示的提示数量上限
MAX_CHAT_HISTORY_MESSAGES = 80   # AI 聊天历史最大消息数
MAX_CHAT_MESSAGE_CHARS = 12000   # 单条 AI 消息最大长度，防止历史长期膨胀
MAX_CONTEXT_TOTAL_CHARS = 40000  # 发给 API 的全部历史消息总字符上限（约 10K token），超出则丢弃最老消息
HOUSEKEEPING_INTERVAL_MS = 5 * 60 * 1000  # 低频运行时清理周期（5分钟）
RESOURCE_SAMPLE_INTERVAL_S = 10 * 60  # 诊断记录中资源趋势的采样间隔
HOUSEKEEPING_GC_EVERY_N = 3  # 每 N 次清理执行一次 gc.collect()
SESSION_SNAPSHOT_INTERVAL_MS = 15000  # 崩溃恢复兜底：定期写入当前会话快照
SESSION_SNAPSHOT_DEBOUNCE_MS = 1200  # 标签/路径变化后的会话快照防抖时间
SESSION_SNAPSHOT_MIN_INTERVAL_MS = 8000  # 事件驱动快照最小间隔，防止 DirPoll/FileWatcher 高频触发写盘
APP_INTERNAL_CHANGE_FILENAMES = {
    'config.json',
    'config.json.tmp',
    'bookmarks.json',
    'chat_history.json',
    'tabex_debug_latest.log',
}

# 大文件夹异步加载配置
LARGE_FOLDER_THRESHOLD = 1000  # 超过此数量文件视为大文件夹
FOLDER_CHECK_TIMEOUT = 500  # 文件夹检查超时时间(ms)
ASYNC_LOAD_ENABLED = True  # 是否启用异步加载
# 慢盘（网络/UNC/映射盘）导航兜底超时：后台解析成功但 NavigateComplete2 因网络中断
# 始终不触发时，用此超时解除“导航中”锁定并隐藏 loading，避免标签永久卡在加载态。
ASYNC_NAV_TIMEOUT_MS = 20000

STATUS_SELECTION_METADATA_LIMIT = 20  # 多选超过阈值时跳过逐项大小统计
# ── 主线程 COM 轮询抗高负载保护 ───────────────────────────────────────────────
# LocationURL 是同步跨进程 COM 调用，无超时。CPU 饱和时其延迟会飙升，阻塞 UI 线程，
# 且卡顿时 Qt 定时器事件堆积、线程一空就爆发式触发，形成无法恢复的“死亡螺旋”。
COM_POLL_SLOW_MS = 180        # 单次 LocationURL 调用超过此耗时即判定系统高负载
COM_POLL_STRESS_BACKOFF_MS = 3000  # 高负载期间将 COM 轮询间隔退避到此值，给 UI 线程喘息
COM_POLL_MIN_GAP_MS = 50      # 挂钟防抖：两次实际轮询的最小真实间隔，吸收卡顿后的爆发触发
COM_POLL_HARD_DEADLINE_MS = 250  # QAx LocationURL 看门狗硬超时：超时即放弃读取并用缓存值
# 状态栏资源占用颜色预警阈值（百分比）：低于 WARN 绿色，WARN~CRIT 橙色，>=CRIT 红色
RESOURCE_WARN_PERCENT = 75
RESOURCE_CRIT_PERCENT = 90
STATUS_UPDATE_DEFER_MS = 80
STATUS_TRACKING_INTERVAL_MS = 220
STATUS_TRACKING_WINDOW_MS = 1400
TITLE_SHORTCUT_EXTENSIONS = ('.lnk', '.exe', '.bat', '.cmd', '.ps1')
SUPPORTED_TERMINAL_TOOLS = ('cmd', 'powershell', 'git-bash')
