from sikuli import *
import time
import sys

# Define image paths
i_yahoo_BBL = r"C:\Users\Kelly\Documents\BBLFF\images\Yahoo_BestBallLobby.jpg"

counter = 0
Settings.WaitScanRate = 0.5  # Set the check interval to 0.5 seconds (500 milliseconds)

def find_image_with_timeout(image, timeout, similarity=0.7, region=None):
    """Attempts to find the image within the specified timeout period and similarity."""
    pattern = Pattern(image).similar(similarity)  # Set similarity for image matching
    start_time = time.time()
    while time.time() - start_time < timeout:
        if region:
            if region.exists(pattern):
                return True
        else:
            if exists(pattern):
                return True
        time.sleep(0.5)  # Sleep to avoid excessive CPU usage
    return False

while True:  # Keep running indefinitely until interrupted
    try:
        # Wait for image_fail with a 10-second timeout  (looking to confirm computer is on YahooBBLFF data/page)
        if not find_image_with_timeout(i_yahoo_BBL, 60): 
            print("Yahoo_BestBallLobby.jpg not found within 60 seconds. Retrying...")
            continue  # Instead of exiting, retry the loop indefinitely

        print("first image found")
        time.sleep(1)
        click(Location(93, 62))      # refresh button
        time.sleep(5)

        hover(Location(400, 283))    # begining of table at Rank
        mouseDown(Button.LEFT)
        hover(Location(860, 600))    # end of table at 0.00
        mouseUp(Button.LEFT)  # (YahooBBLFF data now selected)

        # Clipboard Interference: Add a delay after the Ctrl+C operation to avoid clipboard issues
        type('c', KeyModifier.CTRL)  # (to copy YahooBBLFF selected data)
     
        time.sleep(2)
        # Simulate Alt + Tab (to move to GoogleSheets)
        type(Key.TAB, KeyModifier.ALT)
        
        time.sleep(2)
        # Simulate Ctrl + V   (paste the data in Googlesheets)
        type("v", KeyModifier.CTRL)

        # Simulate Ctrl + Alt + Shift + 5  (initial Googlesheets macro to potentially send emails)
        type("5", KeyModifier.CTRL + KeyModifier.ALT + KeyModifier.SHIFT)
        
        time.sleep(30)  # Allow time for the clipboard operation to complete and for Typebot to move away from Yahoo page.

        time.sleep(2)
        # Simulate Alt + Tab (to move to YahooBBLFF data page)
        type(Key.TAB, KeyModifier.ALT)

        # Example: Stop script based on some condition
        if counter > 100:  # Example condition to stop the script after 100 iterations
            break  # Exit the loop

        counter += 1

    except Exception as e:
        print("An unexpected error occurred: {}".format(e))
        sys.exit(1)
