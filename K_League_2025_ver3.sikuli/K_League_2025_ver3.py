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
IMG_YAHOO       = img_dir + r"\Yahoo.PNG"
IMG_RANK        = img_dir + r"\Rank.PNG"
IMG_END         = img_dir + r"\End.PNG"
IMG_GOOGLESHEET = img_dir + r"\GoogleSheet.PNG"
IMG_INBOX       = img_dir + r"\Inbox.PNG"
IMG_COMPOSE     = img_dir + r"\Compose.PNG"
IMG_BLACKBOX    = img_dir + r"\BlackBox.PNG"

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

# --- Regions (as provided) ---
regionCompose   = Region(4,120,183,110)
regionBlackBox  = Region(685,538,513,360)
regionRank      = Region(491,349,163,156)
regionEnd       = Region(1214,746,266,284)
regionTab       = Region(0,0,1693,47)

# --- Helpers ---
def pat(path):
    return Pattern(path).similar(MATCH_SIM)

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("[{}] {}".format(ts, msg))

def create_exit_flag():
    try:
        with open(exit_flag, "w") as f:
            f.write("exit")
        log("Exit flag created.")
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

def fail_and_request_exit(not_found_name):
    log("{} not found".format(not_found_name))
    short_beep()
    create_exit_flag()

def exists_or_fail(region, image_path, name, timeout=1):
    m = region.exists(pat(image_path), timeout)
    if not m:
        fail_and_request_exit(name)
        return None
    return m

def click_match(m):
    hover(m)
    click(m)

def select_between_matches_and_copy(m1, m2):
    try:
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
    
    matchYahoo = exists_or_fail(regionTab, IMG_YAHOO, "matchYahoo")
    if not matchYahoo:
        return False
    click_match(matchYahoo)
    wait(4)
    type(Key.F5)
    wait(4)
    matchRank = exists_or_fail(regionRank, IMG_RANK, "matchRank")
    if not matchRank:
        return False
    matchEnd = exists_or_fail(regionEnd, IMG_END, "matchEnd")
    if not matchEnd:
        return False
    select_between_matches_and_copy(matchRank, matchEnd)
    wait(3)
    type("c", Key.CTRL)
    if os.path.exists(exit_flag):
        return False
    wait(3)
    matchGoogleSheet = exists_or_fail(regionTab, IMG_GOOGLESHEET, "matchGoogleSheet")
    if not matchGoogleSheet:
        return False
    click_match(matchGoogleSheet)
    wait(5)
    type("v", Key.CTRL)
    wait(3)
    type("1", Key.CTRL + Key.ALT + Key.SHIFT)
    wait(10)
    type("c", Key.CTRL)
    wait(2)
    matchBlackBox = regionBlackBox.exists(pat(IMG_BLACKBOX), 1)
    if not matchBlackBox:
        log("(Email report not sent)")
        wait(1)
        type("3", Key.CTRL + Key.ALT + Key.SHIFT)   
        wait(5)
        return True
    matchInbox = exists_or_fail(regionTab, IMG_INBOX, "matchInbox")
    if not matchInbox:
        short_beep()
        wait(1)
        short_beep()
        return False
    click_match(matchInbox)
    wait(4)
    matchCompose = exists_or_fail(regionCompose, IMG_COMPOSE, "matchCompose")
    if not matchCompose:
        return False
    click_match(matchCompose)
    type("v", Key.CTRL)
    wait(2)
    type(Key.TAB)
    wait(2)
    type("K-League Scoreboard Update")
    wait(2)
    type(Key.TAB)
    wait(2)
    matchGoogleSheet2 = exists_or_fail(regionTab, IMG_GOOGLESHEET, "matchGoogleSheet")
    if not matchGoogleSheet2:
        return False
    click_match(matchGoogleSheet2)
    wait(2)
    type("2", Key.CTRL + Key.ALT + Key.SHIFT)
    wait(5)
    type("c", Key.CTRL)
    wait(2)
    type("3", Key.CTRL + Key.ALT + Key.SHIFT)
    wait(5)
    matchInbox2 = exists_or_fail(regionTab, IMG_INBOX, "matchInbox")
    if not matchInbox2:
        return False
    click_match(matchInbox2)
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

log("Script started.")

while True:
    check_exit_flag_and_quit_if_set()
    now = datetime.now()
    if is_new_quarter(now):
        short_beep()
        log("Quarter-hour trigger at {:02d}:{:02d}".format(now.hour, now.minute))
        do_quarter_hour_actions()
        check_exit_flag_and_quit_if_set()
    wait(1)
