from sikuli import *
import os
import sys
from java.lang import Runtime, Thread
import atexit
from javax.sound.sampled import AudioSystem
from java.io import File
from datetime import datetime

Settings.MoveMouseDelay = 0  # Optional: slows mouse for visibility

# --- Matching threshold (lower = easier/looser) ---
MATCH_SIM = 0.75
Settings.MinSimilarity = MATCH_SIM

# --- Paths ---
exit_flag = r"C:\Users\Kelly\Desktop\exit.flag"
img_dir   = r"C:\Users\Kelly\Documents\BBLFF\images"
custom_wav_path = r"C:\Users\Kelly\Documents\Glacier\custom_beep.wav"

# --- Image files ---
IMG_YAHOO              = img_dir + r"\Yahoo.PNG"
IMG_RANK               = img_dir + r"\Rank.PNG"
IMG_END                = img_dir + r"\End.PNG"
IMG_GOOGLESHEET        = img_dir + r"\GoogleSheet.PNG"
IMG_INBOX              = img_dir + r"\Inbox.PNG"
IMG_COMPOSE            = img_dir + r"\Compose.PNG"
IMG_FOOTBALL           = img_dir + r"\Football.PNG"
IMG_REFRESH_SUCCESS    = img_dir + r"\Refresh_Success.PNG"

# --- Brightness Reset on Exit ---
def set_brightness(level):
    try:
        cmd = 'powershell.exe -Command "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{})"'.format(level)
        Runtime.getRuntime().exec(cmd)
    except:
        print("Failed to set brightness")

atexit.register(lambda: set_brightness(100))

# --- Sound Playback ---
def play_sound(path_to_wav):
    try:
        sound_file = File(path_to_wav)
        if not sound_file.exists():
            print("Sound file not found:", path_to_wav)
            return
        audio_in = AudioSystem.getAudioInputStream(sound_file)
        clip = AudioSystem.getClip()
        clip.open(audio_in)
        clip.start()
        Thread.sleep(clip.getMicrosecondLength() / 1000)
        clip.stop()
        clip.close()
    except Exception as e:
        print("Sound error:", e)

def short_beep():
    play_sound(custom_wav_path)

# --- Regions (as provided, with Football rename) ---
regionCompose   = Region(4,120,183,110)
regionFootball  = Region(601,501,573,441)  # renamed from regionBlackBox
regionRank      = Region(491,349,163,156)
regionEnd       = Region(1214,746,266,284)
regionTab       = Region(0,0,1693,47)
regionRefreshSuccess = Region(474,105,470,250)

# --- Helpers ---
def pat(path, sim=None):
    """Return a Pattern for 'path' with either provided 'sim' or the global MATCH_SIM."""
    p = Pattern(path)
    if sim is None:
        return p.similar(MATCH_SIM)
    return p.similar(sim)

def region_str(r):
    return "Region(x={}, y={}, w={}, h={})".format(r.getX(), r.getY(), r.getW(), r.getH())

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("[{}] {}".format(ts, msg))

def create_exit_flag():
    try:
        with open(exit_flag, "w") as f:
            f.write("exit")
        log("Exit flag created.")
        short_beep()
    except Exception as e:
        log("Failed to create exit flag: {}".format(e))

def check_exit_flag_and_quit_if_set():
    if os.path.exists(exit_flag):
        try:
            os.remove(exit_flag)
        except:
            pass
        log("Exit flag detected. Exiting script.")
        set_brightness(100)
        sys.exit(0)

def fail_and_request_exit(not_found_name, image_path=None, region=None, timeout=None, sim=None):
    pieces = ["{} not found".format(not_found_name)]
    if image_path: pieces.append("img='{}'".format(image_path))
    if sim is not None: pieces.append("similarity={:.2f}".format(sim))
    if region: pieces.append(region_str(region))
    if timeout is not None: pieces.append("timeout={}s".format(timeout))
    log(" | ".join(pieces))
    short_beep()
    create_exit_flag()

def exists_or_fail(region, image_path, name, timeout=1, sim=None):
    """Try to find image in region; log rich context on failure and return Match or None."""
    log("Searching '{}' in {} with sim={} timeout={}s".format(name, region_str(region), 
                                                             "{:.2f}".format(sim if sim is not None else MATCH_SIM), timeout))
    try:
        m = region.exists(pat(image_path, sim), timeout)
    except FindFailed as e:
        log("FindFailed on '{}': {}".format(name, e))
        m = None
    if not m:
        fail_and_request_exit(name, image_path=image_path, region=region, timeout=timeout, sim=(sim if sim is not None else MATCH_SIM))
        return None
    log("Found '{}' at ({}, {}) size=({}x{})".format(name, m.getX(), m.getY(), m.getW(), m.getH()))
    return m

def click_match(m):
    hover(m)
    click(m)

def select_between_matches_and_copy(m1, m2):
    try:
        log("Drag-select from ({},{}) to ({},{}) then copy".format(m1.getTarget().getX(), m1.getTarget().getY(),
                                                                   m2.getTarget().getX(), m2.getTarget().getY()))
        dragDrop(m1.getTarget(), m2.getTarget())
        type("c", Key.CTRL)
    except Exception as e:
        log("Selection/Copy failed: {}".format(e))
        fail_and_request_exit("DragSelectCopy")

def wake_display():
    try:
        curr = Env.getMouseLocation()
    except:
        curr = Location(10, 10)
    hover(Location(curr.x + 1, curr.y + 1))
    wait(0.4)
    hover(curr)
    wait(0.6)

# --- Quarter-hour trigger control ---
last_trigger_hm = None

def is_new_quarter(now_dt):
    global last_trigger_hm
    if now_dt.minute % 15 != 0:
        return False
    hm_tuple = (now_dt.year, now_dt.month, now_dt.day, now_dt.hour, now_dt.minute)
    if hm_tuple == last_trigger_hm:
        return False
    last_trigger_hm = hm_tuple
    return True

# --- Quarter-hour routine ---
def do_quarter_hour_actions():
    wake_display()
    wait(4)
    
    matchYahoo = exists_or_fail(regionTab, IMG_YAHOO, "Yahoo Tab", timeout=3)
    if not matchYahoo:
        return False
    click_match(matchYahoo)
    wait(4)
    type(Key.F5)  # First refresh
    wait(4)  # Give page time to load after first refresh

    matchRefreshSuccess = regionRefreshSuccess.exists(pat(IMG_REFRESH_SUCCESS), 5)
    if not matchRefreshSuccess:
        log("Refresh_Success not found after first refresh. Retrying...")
        type(Key.F5)  # Second refresh
        wait(10)  # Longer wait to ensure page reloads fully
        matchRefreshSuccess = regionRefreshSuccess.exists(pat(IMG_REFRESH_SUCCESS), 5)
        if not matchRefreshSuccess:
            log("Yahoo page refresh failed, exited program.")
            create_exit_flag()        # <-- create C:\Users\Kelly\Desktop\exit.flag
            set_brightness(100)       # restore brightness (matches other exits)
            sys.exit(0)

    log("Refresh_Success image found. Proceeding...")

    wait(1)
    matchRank = exists_or_fail(regionRank, IMG_RANK, "Rank Anchor", timeout=3)
    if not matchRank:
        return False

    matchEnd = exists_or_fail(regionEnd, IMG_END, "End Anchor", timeout=3)
    if not matchEnd:
        return False

    select_between_matches_and_copy(matchRank, matchEnd)
    wait(3)
    type("c", Key.CTRL)
    if os.path.exists(exit_flag):
        return False

    wait(3)
    matchGoogleSheet = exists_or_fail(regionTab, IMG_GOOGLESHEET, "Google Sheet Tab", timeout=3)
    if not matchGoogleSheet:
        return False

    click_match(matchGoogleSheet)
    wait(5)
    type("v", Key.CTRL)
    wait(3)
    type("1", Key.CTRL + Key.ALT + Key.SHIFT)
    wait(17)
    type("c", Key.CTRL)
    wait(2)

    # Football check (was BlackBox), explicitly allow looser match
    matchFootball = regionFootball.exists(Pattern(IMG_FOOTBALL).similar(0.40), 5)
    if matchFootball:
        log("Found 'Football' within {} at ({}, {})".format(region_str(regionFootball), matchFootball.getX(), matchFootball.getY()))
    else:
        log("Football NOT found in {} (img='{}', similarity=0.40, timeout=5s)".format(region_str(regionFootball), IMG_FOOTBALL))
        log("(Email report not sent) -> Continuing with reset actions")
        wait(1)
        type("3", Key.CTRL + Key.ALT + Key.SHIFT)
        wait(5)
        return True

    # Inbox: lower similarity to 0.40 based on visual variance
    matchInbox = exists_or_fail(regionTab, IMG_INBOX, "Gmail Inbox Tab", timeout=3)
    if not matchInbox:
        return False              
    click_match(matchInbox)
    wait(4)
    matchCompose = exists_or_fail(regionCompose, IMG_COMPOSE, "Compose Button", timeout=5, sim=0.60)
    if not matchCompose:
        return False
    click_match(matchCompose)
    wait(3)
    type("v", Key.CTRL)
    wait(2)
    type(Key.TAB)
    wait(2)
    type("K-League Scoreboard Update")
    wait(2)
    type(Key.TAB)
    wait(2)

    matchGoogleSheet2 = exists_or_fail(regionTab, IMG_GOOGLESHEET, "Google Sheet Tab (2nd)", timeout=3)
    if not matchGoogleSheet2:
        return False
    click_match(matchGoogleSheet2)
    wait(4)
    type("2", Key.CTRL + Key.ALT + Key.SHIFT)
    wait(5)
    type("c", Key.CTRL)
    wait(2)
    type("3", Key.CTRL + Key.ALT + Key.SHIFT)
    wait(5)

    matchInbox2 = exists_or_fail(regionTab, IMG_INBOX, "Gmail Inbox Tab (Return)", timeout=3)
    if not matchInbox2:
        return False
    click_match(matchInbox2)
    wait(4)
    type("v", Key.CTRL)
    wait(2)
    type(Key.TAB)
    wait(2)
    type(Key.SPACE)
    wait(2)
    return True

# --- Script Start ---
set_brightness(100)
wait(2)

log("Script started. MATCH_SIM={:.2f}, regions: Compose={}, Football={}, Rank={}, End={}, Tab={}".format(
    MATCH_SIM, region_str(regionCompose), region_str(regionFootball), region_str(regionRank), region_str(regionEnd), region_str(regionTab)
))

while True:
    check_exit_flag_and_quit_if_set()
    now = datetime.now()
    if is_new_quarter(now):
        short_beep()
        log("Quarter-hour trigger at {:02d}:{:02d}".format(now.hour, now.minute))
        do_quarter_hour_actions()
        check_exit_flag_and_quit_if_set()
    wait(1)
