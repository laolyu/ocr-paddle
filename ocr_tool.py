"""
OCR 截图识别工具
快捷键: Alt+Q → 框选区域 → 自动识别并复制到剪贴板
- 智能模式：自动判断纯文本 / 表格，表格输出 TSV（粘贴到 Excel 自动分列）
"""
import ctypes
import re
import sys
import threading
import warnings
import logging
from queue import Queue

# 1. 【防御性设置】强制 matplotlib 使用无头后端，防止 PaddleOCR 底层意外唤起 TkAgg 冲突
import matplotlib
matplotlib.use('Agg') 

# 2. 配置日志与警告
logging.basicConfig(
    level=logging.ERROR,  # DEBUG 级别以便捕获所有信息
    format='%(asctime)s.%(msecs)03d [%(levelname)s][%(threadName)s] %(message)s',
    datefmt='%H:%M:%S'
)
warnings.filterwarnings('ignore', module='requests')
logging.getLogger("ppocr").setLevel(logging.WARNING)
log = logging.getLogger('ocr_tool')

# 3. Windows 高分屏 DPI 适配（必须在 tkinter 之前调用）
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# 4. 【关键修复：全局唯一隐藏根窗口】
import tkinter as tk
GLOBAL_ROOT = tk.Tk()
GLOBAL_ROOT.withdraw() # 永远隐藏，只作为事件循环的宿主

import numpy as np
import pyperclip
import mss
from pynput import mouse as pmouse  # noqa: F401
from pynput import keyboard as pkeyboard
from PIL import Image, ImageTk
from paddleocr import PaddleOCR

HOTKEY = 'alt+q'

print("正在加载 OCR 模型（首次运行会下载模型文件，请稍候...）")
ocr = PaddleOCR(
    use_doc_orientation_classify=False,
    use_doc_unwarping=False,
    use_textline_orientation=False,
    device="gpu", # 如果你有N卡并装了GPU版PyTorch，可以改为 "gpu"
    engine="transformers",
)
print(f"✓ 就绪！按 {HOTKEY.upper()} 框选识别  |  Esc 取消  |  Ctrl+C 退出\n")
log.debug("OCR 引擎初始化完成，开始监听快捷键")


# ───────────────────────── 框选 UI ─────────────────────────
_VK_LBUTTON = 0x01
_user32 = ctypes.windll.user32

class _POINT(ctypes.Structure):
    _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]

def select_region(full_img: Image.Image):
    log.debug("select_region: 进入，等待用户框选")
    sel = {'coords': None}
    state = {'phase': 'wait', 'sx': 0, 'sy': 0}

    # 使用 Toplevel 挂载到全局 ROOT 上
    top = tk.Toplevel(GLOBAL_ROOT)
    top.withdraw()
    top.overrideredirect(True)
    top.attributes('-topmost', True)
    top.config(bg='black', cursor='crosshair')

    w, h = full_img.size
    top.geometry(f'{w}x{h}+0+0')

    photo = ImageTk.PhotoImage(full_img)
    cvs = tk.Canvas(top, width=w, height=h, highlightthickness=0, bg='black', cursor='crosshair')
    cvs.pack()
    cvs.create_image(0, 0, anchor='nw', image=photo)

    ov_full = cvs.create_rectangle(0, 0, w, h, fill='black', stipple='gray25', outline='')
    hint_id = cvs.create_text(w // 2, 36, text='框选要识别的区域    Esc 取消', fill='white', font=('Microsoft YaHei', 16))

    ov_top = cvs.create_rectangle(0, 0, w, 0, fill='black', stipple='gray50', outline='', state='hidden')
    ov_bot = cvs.create_rectangle(0, h, w, h, fill='black', stipple='gray50', outline='', state='hidden')
    ov_lft = cvs.create_rectangle(0, 0, 0, h, fill='black', stipple='gray50', outline='', state='hidden')
    ov_rgt = cvs.create_rectangle(w, 0, w, h, fill='black', stipple='gray50', outline='', state='hidden')
    sel_border = cvs.create_rectangle(0, 0, 0, 0, outline='white', width=2, dash=(6, 4), state='hidden')

    after_id = [None]

    def _cancel(_=None):
        log.debug("select_region: 收到取消信号 (ESC/Escape)")
        state['phase'] = 'cancel'

    # 注意：这里【不使用】keyboard.on_press_key，因为其回调在 keyboard 库的钩子线程中执行，
    # 若随后在主线程调用 keyboard.unhook，会产生竞态条件破坏 keyboard 库内部钩子列表，
    # 导致 alt+q 热键永久失效。覆盖层已 focus_force，tkinter 绑定足够捕获 Escape。
    top.bind('<Escape>', _cancel)
    cvs.bind('<Escape>', _cancel)  # canvas 也绑定，防止焦点在 canvas 上时漏掉

    def _poll():
        phase = state['phase']

        if phase == 'cancel':
            log.debug("_poll: 进入 cancel 分支，准备退出 mainloop")
            if after_id[0]:
                GLOBAL_ROOT.after_cancel(after_id[0])
                after_id[0] = None
            try:
                top.destroy()
            except Exception as e:
                log.warning(f"_poll cancel: top.destroy() 异常: {e}")
            GLOBAL_ROOT.quit()
            log.debug("_poll: mainloop quit 已调用 (cancel)")
            return

        lmb_down = bool(_user32.GetAsyncKeyState(_VK_LBUTTON) & 0x8000)
        pt = _POINT()
        _user32.GetCursorPos(ctypes.byref(pt))
        mx, my = pt.x, pt.y

        if phase == 'wait':
            if lmb_down:
                state['phase'] = 'drag'
                state['sx'], state['sy'] = mx, my
                cvs.itemconfig(ov_full, state='hidden')
                cvs.itemconfig(hint_id, state='hidden')
                for it in (ov_top, ov_bot, ov_lft, ov_rgt, sel_border):
                    cvs.itemconfig(it, state='normal')

        elif phase == 'drag':
            sx, sy = state['sx'], state['sy']
            x1, y1 = min(sx, mx), min(sy, my)
            x2, y2 = max(sx, mx), max(sy, my)
            cvs.coords(ov_top, 0,  0,  w,  y1)
            cvs.coords(ov_bot, 0,  y2, w,  h)
            cvs.coords(ov_lft, 0,  y1, x1, y2)
            cvs.coords(ov_rgt, x2, y1, w,  y2)
            cvs.coords(sel_border, x1, y1, x2, y2)

            if not lmb_down:
                if x2 - x1 > 5 and y2 - y1 > 5:
                    sel['coords'] = (x1, y1, x2, y2)
                    log.debug(f"_poll: 框选完成 coords={sel['coords']}")
                else:
                    log.debug("_poll: 框选区域过小，取消")
                state['phase'] = 'done'
                if after_id[0]:
                    GLOBAL_ROOT.after_cancel(after_id[0])
                    after_id[0] = None
                try:
                    top.destroy()
                except Exception as e:
                    log.warning(f"_poll done: top.destroy() 异常: {e}")
                GLOBAL_ROOT.quit()
                log.debug("_poll: mainloop quit 已调用 (done)")
                return

        after_id[0] = GLOBAL_ROOT.after(16, _poll)

    top.update_idletasks()
    top.deiconify()
    top.lift()
    top.focus_force()
    after_id[0] = GLOBAL_ROOT.after(16, _poll)
    log.debug("select_region: 进入 mainloop (等待框选)")
    GLOBAL_ROOT.mainloop()
    log.debug(f"select_region: mainloop 已退出，coords={sel['coords']}")
    return sel['coords']


# ───────────────────────── 走马灯边框 ─────────────────────────
class MarchingAntsBorder:
    COLORS = ['#FF4081', '#FFD740', '#69F0AE', '#40C4FF', '#B388FF', '#FF6E40']

    def __init__(self, x1, y1, x2, y2, thickness=4):
        self.x1, self.y1, self.x2, self.y2 = x1, y1, x2, y2
        self.t = thickness
        self.top = None
        self._stop = False
        self._frame = 0
        self._after_id = None

    def show(self):
        self.top = tk.Toplevel(GLOBAL_ROOT)
        self.top.withdraw()
        self.top.overrideredirect(True)
        self.top.attributes('-topmost', True)
        self.top.config(bg='magenta')
        try:
            self.top.attributes('-transparentcolor', 'magenta')
        except tk.TclError:
            pass
        sw = self.top.winfo_screenwidth()
        sh = self.top.winfo_screenheight()
        self.top.geometry(f'{sw}x{sh}+0+0')
        self.cvs = tk.Canvas(self.top, width=sw, height=sh, highlightthickness=0, bg='magenta')
        self.cvs.pack()
        self._draw()
        self.top.deiconify()
        self.top.lift()
        self._tick()

    def _draw(self):
        self.cvs.delete('ants')
        x1, y1, x2, y2 = self.x1, self.y1, self.x2, self.y2
        seg, gap = 12, 6
        total = seg + gap
        offset = (self._frame * 3) % total
        
        def draw_line(a, b, c, d, horizontal):
            length = (c - a) if horizontal else (d - b)
            pos = -offset
            i = 0
            while pos < length:
                color = self.COLORS[(i + self._frame // 2) % len(self.COLORS)]
                p1, p2 = max(0, pos), min(length, pos + seg)
                if p2 > p1:
                    if horizontal: self.cvs.create_line(a+p1, b, a+p2, d, fill=color, width=self.t, tags='ants', capstyle='round')
                    else: self.cvs.create_line(a, b+p1, c, b+p2, fill=color, width=self.t, tags='ants', capstyle='round')
                pos += total
                i += 1
        draw_line(x1, y1, x2, y1, True)
        draw_line(x1, y2, x2, y2, True)
        draw_line(x1, y1, x1, y2, False)
        draw_line(x2, y1, x2, y2, False)

    def _tick(self):
        if self._stop or not self.top:
            return
        self._frame += 1
        try:
            self._draw()
            self._after_id = GLOBAL_ROOT.after(60, self._tick)
        except tk.TclError:
            pass

    def close(self):
        log.debug("MarchingAntsBorder.close: 开始清理")
        self._stop = True
        if self.top:
            # 先取消 after 回调，再销毁窗口，最后退出 mainloop
            # quit() 必须在 destroy() 之前且保证被调用
            try:
                if self._after_id:
                    GLOBAL_ROOT.after_cancel(self._after_id)
                    self._after_id = None
            except Exception as e:
                log.warning(f"MarchingAntsBorder.close: after_cancel 异常: {e}")
            try:
                self.top.destroy()
            except Exception as e:
                log.warning(f"MarchingAntsBorder.close: top.destroy() 异常: {e}")
            finally:
                self.top = None
            # quit() 放在最外层确保一定被调用
            try:
                GLOBAL_ROOT.quit()
                log.debug("MarchingAntsBorder.close: mainloop quit 已调用")
            except Exception as e:
                log.error(f"MarchingAntsBorder.close: GLOBAL_ROOT.quit() 异常: {e}")
        else:
            log.debug("MarchingAntsBorder.close: top 已为 None，跳过 quit")


# ───────────────────────── 表格重建 ─────────────────────────
def _extract_cells(result):
    cells = []
    for res in result:
        texts = res.get('rec_texts', [])
        boxes = res.get('rec_boxes', None)
        if boxes is None or len(boxes) == 0:
            polys = res.get('rec_polys', [])
            for t, p in zip(texts, polys):
                if not t or not t.strip(): continue
                p = np.asarray(p)
                cells.append((t.strip(), int(p[:, 0].min()), int(p[:, 1].min()), int(p[:, 0].max()), int(p[:, 1].max())))
            continue
        for t, b in zip(texts, boxes):
            if not t or not t.strip(): continue
            x1, y1, x2, y2 = (int(v) for v in b[:4])
            cells.append((t.strip(), x1, y1, x2, y2))
    return cells

def _cluster_1d(values, tolerance):
    if not values: return []
    order = sorted(range(len(values)), key=lambda i: values[i])
    cluster_ids = [0] * len(values)
    cur_id = 0
    cluster_ids[order[0]] = cur_id
    last_v = values[order[0]]
    for idx in order[1:]:
        v = values[idx]
        if v - last_v > tolerance: cur_id += 1
        cluster_ids[idx] = cur_id
        last_v = v
    return cluster_ids

def _build_table(cells):
    if not cells: return [], False
    heights = [y2 - y1 for _, _, y1, _, y2 in cells]
    avg_h = max(1, sum(heights) / len(heights))
    y_centers = [(y1 + y2) / 2 for _, _, y1, _, y2 in cells]
    row_ids = _cluster_1d(y_centers, avg_h * 0.6)

    by_row = {}
    for c, rid in zip(cells, row_ids): by_row.setdefault(rid, []).append(c)
    sorted_row_ids = sorted(by_row, key=lambda r: np.mean([(c[2] + c[4]) / 2 for c in by_row[r]]))
    row_cells = [sorted(by_row[r], key=lambda c: c[1]) for r in sorted_row_ids]

    multi_col_rows = sum(1 for r in row_cells if len(r) >= 2)
    is_table = len(row_cells) >= 2 and multi_col_rows >= max(2, len(row_cells) // 2)

    if not is_table:
        rows = [[' '.join(c[0] for c in r)] for r in row_cells]
        return rows, False

    widths = [c[3] - c[1] for r in row_cells for c in r]
    avg_w = max(1, sum(widths) / len(widths))
    flat_x = [(c[1] + c[3]) / 2 for r in row_cells for c in r]
    col_ids_flat = _cluster_1d(flat_x, avg_w * 0.7)
    n_cols = max(col_ids_flat) + 1 if col_ids_flat else 1

    cursor = 0
    rows = []
    for r in row_cells:
        row_arr = [''] * n_cols
        for c in r:
            cid = col_ids_flat[cursor]
            cursor += 1
            row_arr[cid] = (row_arr[cid] + ' ' + c[0]).strip() if row_arr[cid] else c[0]
        rows.append(row_arr)
    return rows, True

def _format_tsv(rows):
    return '\n'.join('\t'.join(cells) for cells in rows)


# ───────────────────────── OCR 数字混淆自动纠错 ─────────────────────────
_DIGIT_CONFUSION_MAP = str.maketrans({
    # → 1
    'H': '1', 'I': '1', 'i': '1', 'l': '1', '|': '1', '\u5de5': '1',  # 工
    # → 0
    'O': '0', 'o': '0', 'D': '0', 'e': '0', '\u65e5': '0',  # 日
    # → 其他数字
    'B': '8',
    'S': '5', 's': '5',
    'Z': '2', 'z': '2',
    'G': '6', 'g': '9', 'q': '9',
})
_NUMBER_RE = re.compile(r'^[+\-]?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?%?$')
# 用于列检测：匹配以数字或混淆字符开头的单元格（含「100e（注释）」这类复合格）
_NUM_START_RE = re.compile(r'^[0-9\u5de5\u65e5HeIilOoDeSsZzGgqB|]')

# 匹配含至少一个真实数字或「工/日」的混淆字符串，且不被字母/中文包围
# CJK 边界排除 工(U+5DE5) 和 日(U+65E5)，因为它们是已列入混淆字符的汉字
_MIXED_NUM_RE = re.compile(
    '(?<![a-zA-Z\u4e00-\u5de4\u5de6-\u65e4\u65e6-\u9fff])'
    '(?=[0-9HeIilOoDeSsZzGgqB|\u5de5\u65e5]*[0-9\u5de5\u65e5])'  # 含真实数字或工/日
    '([0-9HeIilOoDeSsZzGgqB|\u5de5\u65e5]+)'
    '(?![a-zA-Z\u4e00-\u5de4\u5de6-\u65e4\u65e6-\u9fff])'
)

def _is_number(text):
    return bool(_NUMBER_RE.match(text.strip()))

def _try_fix_number(text):
    """整个字符串替换混淆字符后若为合法数字则返回纠正值，否则 None。"""
    candidate = text.strip().translate(_DIGIT_CONFUSION_MAP)
    return candidate if _is_number(candidate) else None

def _fix_in_text(text):
    """在混合文本中查找含真实数字/工的混淆串并纠正（保留其余文字不变）。"""
    def _fix_run(m):
        token = m.group(1)
        fixed = _try_fix_number(token)
        if fixed is not None and fixed != token:
            log.info(f'自动纠错: {repr(token)} \u2192 {repr(fixed)}')
            return fixed
        return token
    return _MIXED_NUM_RE.sub(_fix_run, text)

def _auto_correct(rows, is_table):
    """
    纠错策略：
    • 表格模式 —— 逐列检测「数字列」（≥70% 单元格是/可纠为合法数字）：
        - 纯数字单元格：整体替换（含无真实数字的混淆串，如 He→10）
        - 复合单元格：  定位内嵌数字片段后局部替换（如 100e(…)→1000(…)）
    • 文本模式 —— 对每个单元格做局部替换（要求含真实数字/工，防止误改英文）
    """
    if not rows:
        return rows

    if not is_table:
        return [[_fix_in_text(cell) for cell in row] for row in rows]

    # 预计算每格的整体纠错候选
    candidates = [[_try_fix_number(cell) for cell in row] for row in rows]
    n_cols = max((len(r) for r in rows), default=0)

    # 逐列判断是否为数字列
    # 计数条件：整格是合法数字 / 可整体纠为合法数字 / 以数字或混淆字符开头（含注释的复合格）
    numeric_col = [False] * n_cols
    for col in range(n_cols):
        total = num = 0
        for r in range(len(rows)):
            if col >= len(rows[r]) or not rows[r][col].strip():
                continue
            total += 1
            cell = rows[r][col]
            if (_is_number(cell)
                    or (col < len(candidates[r]) and candidates[r][col] is not None)
                    or bool(_NUM_START_RE.match(cell.strip()))):
                num += 1
        if total >= 1 and num / total >= 0.7:
            numeric_col[col] = True

    result = []
    for r, row in enumerate(rows):
        new_row = []
        for col, cell in enumerate(row):
            if col < n_cols and numeric_col[col]:
                cand = candidates[r][col] if col < len(candidates[r]) else None
                if cand is not None and cand != cell.strip():
                    # 整体可纠（含 He→10 这类无真实数字的混淆串）
                    log.info(f'自动纠错: {repr(cell)} \u2192 {repr(cand)}')
                    new_row.append(cand)
                else:
                    # 复合单元格：局部修复内嵌数字片段
                    new_row.append(_fix_in_text(cell))
            else:
                # 非数字列也做局部修复（含真实数字的混淆串，如 1eee→1000）
                new_row.append(_fix_in_text(cell))
        result.append(new_row)
    return result


# ───────────────────────── 主流程 ─────────────────────────
_run_count = 0  # 记录调用次数，便于日志追踪

def run_ocr():
    global _run_count
    _run_count += 1
    seq = _run_count
    log.info(f"=== [#{seq}] 开始新的截图识别流程 ===")
    log.debug(f"[#{seq}] 当前活跃线程数: {threading.active_count()}")

    try:
        with mss.MSS() as sct:
            shot = sct.grab(sct.monitors[1])
            full_img = Image.frombytes('RGB', shot.size, shot.bgra, 'raw', 'BGRX')
        log.debug(f"[#{seq}] 截图完成，尺寸={full_img.size}")
    except Exception as e:
        log.error(f"[#{seq}] 截图失败: {e}")
        return

    coords = select_region(full_img)
    if not coords:
        log.info(f"[#{seq}] 流程取消：未选择区域")
        return
    log.info(f"[#{seq}] 已框选区域: {coords}")

    cropped = full_img.crop(coords)

    ants = MarchingAntsBorder(*coords, thickness=4)
    try:
        ants.show()
        log.debug(f"[#{seq}] 走马灯边框已显示")
    except Exception as e:
        log.error(f"[#{seq}] 走马灯边框显示失败: {e}")
        return

    holder = {'done': False}
    ocr_start = threading.Event()

    def _worker():
        ocr_start.set()
        log.debug(f"[#{seq}] WORKER: OCR 线程已启动")
        try:
            holder['result'] = ocr.predict(np.array(cropped))
            log.debug(f"[#{seq}] WORKER: OCR 完成")
        except Exception as exc:
            log.error(f"[#{seq}] WORKER: OCR 抛出异常: {exc}", exc_info=True)
            holder['error'] = exc
        finally:
            holder['done'] = True
            log.debug(f"[#{seq}] WORKER: done 标志已设置")

    t = threading.Thread(target=_worker, name=f"OCR-Worker-{seq}", daemon=True)
    t.start()
    log.debug(f"[#{seq}] OCR 工作线程已启动: {t.name}")

    _check_count = [0]
    _OCR_TIMEOUT_MS = 60_000  # 最长等待 60 秒

    def _check_done():
        _check_count[0] += 1
        elapsed_ms = _check_count[0] * 50
        if holder.get('done'):
            log.debug(f"[#{seq}] _check_done: OCR 已完成 (elapsed≈{elapsed_ms}ms)，关闭边框")
            try:
                ants.close()
            except Exception as e:
                log.error(f"[#{seq}] _check_done: ants.close() 异常，强制 quit: {e}", exc_info=True)
                try:
                    GLOBAL_ROOT.quit()
                except Exception:
                    pass
        elif elapsed_ms >= _OCR_TIMEOUT_MS:
            log.error(f"[#{seq}] _check_done: OCR 超时 ({_OCR_TIMEOUT_MS}ms)，强制退出 mainloop！")
            try:
                ants.close()
            except Exception:
                pass
            try:
                GLOBAL_ROOT.quit()
            except Exception:
                pass
        else:
            if _check_count[0] % 20 == 0:  # 每秒打印一次等待日志
                log.debug(f"[#{seq}] _check_done: 等待 OCR 完成... elapsed≈{elapsed_ms}ms")
            GLOBAL_ROOT.after(50, _check_done)

    GLOBAL_ROOT.after(50, _check_done)
    log.debug(f"[#{seq}] 进入 mainloop (等待 OCR)")
    GLOBAL_ROOT.mainloop()
    log.debug(f"[#{seq}] mainloop 已退出")

    if not holder.get('done'):
        log.error(f"[#{seq}] mainloop 退出但 OCR 未完成，异常状态！")

    if 'error' in holder:
        log.error(f"[#{seq}] × 识别异常: {holder['error']}")
        return

    cells = _extract_cells(holder.get('result', []))
    if not cells:
        log.warning(f"[#{seq}] × 未检测到文字")
        return

    rows, is_table = _build_table(cells)
    rows = _auto_correct(rows, is_table)
    output = _format_tsv(rows)
    pyperclip.copy(output)

    tag = '[表格]' if is_table else '[文本]'
    n_rows = len(rows)
    n_cols = max((len(r) for r in rows), default=0)
    log.info(f"[#{seq}] ✓ 已复制 {tag} {n_rows}行 x {n_cols}列")


# ───────────────────────── 程序入口 ─────────────────────────

trigger = Queue()

def _hotkey_fired():
    qsize = trigger.qsize()
    log.info(f"[HOTKEY HIT] 快捷键触发！当前队列中已有 {qsize} 个待处理请求")
    if qsize >= 2:
        log.warning(f"队列积压 {qsize} 个请求，上一次识别可能仍在进行中")
    trigger.put(1)

# 用 pynput.HotKey 替代 keyboard.add_hotkey，解决 Alt 修饰键组合不触发问题
_hotkey_combo = pkeyboard.HotKey(
    pkeyboard.HotKey.parse('<alt>+q'),
    _hotkey_fired,
)

def _pk_on_press(key):
    try:
        _hotkey_combo.press(_pk_listener.canonical(key))
    except Exception as e:
        log.debug(f"pynput on_press 异常: {e}")

def _pk_on_release(key):
    try:
        _hotkey_combo.release(_pk_listener.canonical(key))
    except Exception as e:
        log.debug(f"pynput on_release 异常: {e}")

_pk_listener = pkeyboard.Listener(on_press=_pk_on_press, on_release=_pk_on_release)
_pk_listener.start()
log.info(f"快捷键 {HOTKEY.upper()} 已注册 (pynput)")
log.debug(f"pynput 监听线程: {_pk_listener} alive={_pk_listener.is_alive()}")

try:
    loop_count = 0
    while True:
        log.debug(f"主循环等待快捷键触发... (已完成 {loop_count} 次)")
        trigger.get()
        loop_count += 1
        log.debug(f"主循环: 收到第 {loop_count} 次触发，调用 run_ocr()")
        try:
            run_ocr()
        except Exception as e:
            log.error(f"主循环: run_ocr() 意外抛出异常: {e}", exc_info=True)
        log.debug(f"主循环: run_ocr() 已返回，回到等待状态")
except KeyboardInterrupt:
    print("\n已退出")
    sys.exit(0)