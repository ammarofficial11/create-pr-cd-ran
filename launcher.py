import os
import sys
import threading
import time
import webbrowser
import uvicorn

exe_dir = os.path.dirname(
    os.path.abspath(sys.argv[0])
)

os.chdir(exe_dir)

sys.path.insert(0, exe_dir)

from api.app import app


def start_server():

    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8000
    )


def main():

    threading.Thread(
        target=start_server,
        daemon=True
    ).start()

    time.sleep(3)

    webbrowser.open(
        "http://127.0.0.1:8000/ui"
    )

    while True:
        time.sleep(1)


if __name__ == "__main__":
    main()