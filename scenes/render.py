"""Render scenes.html to PNG frames via Playwright + Edge."""
import sys, os, pathlib
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).parent
FRAMES = ROOT / "frames"
FPS = 30

def main():
    t0 = float(sys.argv[1]) if len(sys.argv) > 1 else 0.0
    t1 = float(sys.argv[2]) if len(sys.argv) > 2 else 30.0
    out = pathlib.Path(sys.argv[3]) if len(sys.argv) > 3 else FRAMES
    keyframes = sys.argv[4] == "keys" if len(sys.argv) > 4 else False
    out.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=True,
            args=["--force-device-scale-factor=1", "--hide-scrollbars"])
        page = browser.new_page(viewport={"width": 1080, "height": 1920})
        page.goto((ROOT / "scenes.html").as_uri())
        page.evaluate("window.READY")
        page.wait_for_timeout(500)

        if keyframes:
            times = [float(x) for x in sys.argv[5].split(",")]
            for t in times:
                page.evaluate(f"seek({t})")
                page.wait_for_timeout(40)
                page.screenshot(path=str(out / f"key_{t:05.2f}.png"))
                print(f"key {t}")
        else:
            f0, f1 = int(t0 * FPS), int(t1 * FPS)
            for f in range(f0, f1):
                t = f / FPS
                page.evaluate(f"seek({t})")
                page.evaluate("(async()=>{const i=document.getElementById('s4img'); if(i){try{await i.decode();}catch(e){}}})()")
                page.screenshot(path=str(out / f"f{f:05d}.png"))
                if f % 60 == 0:
                    print(f"frame {f}/{f1}", flush=True)
        browser.close()
    print("DONE")

if __name__ == "__main__":
    main()
