import json
import logging
import os

# Setup logging
logging.basicConfig(filename='data/logs/app.log', level=logging.INFO)

def save_cookies(driver, path):
    with open(path, 'w') as f:
        json.dump(driver.get_cookies(), f)
    logging.info("Cookies saved to %s", path)

def load_cookies(driver, path):
    if os.path.exists(path):
        with open(path, 'r') as f:
            cookies = json.load(f)
        for cookie in cookies:
            driver.add_cookie(cookie)
        logging.info("Cookies loaded from %s", path)

def save_local_storage(driver, path):
    # Extract all localStorage key/values from the current origin
    storage = driver.execute_script(
        """
        var ls = window.localStorage;
        var items = {};
        for (var i = 0; i < ls.length; i++) {
            var k = ls.key(i);
            items[k] = ls.getItem(k);
        }
        return items;
        """
    )
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(storage or {}, f, ensure_ascii=False)
    logging.info("LocalStorage saved to %s", path)

def load_local_storage(driver, path):
    if not os.path.exists(path):
        return
    with open(path, 'r', encoding='utf-8') as f:
        try:
            items = json.load(f)
        except Exception:
            items = {}
    if isinstance(items, dict):
        # Set each key back into localStorage for current origin
        for k, v in items.items():
            try:
                driver.execute_script(
                    "window.localStorage.setItem(arguments[0], arguments[1]);",
                    k,
                    v if v is not None else ""
                )
            except Exception:
                continue
        logging.info("LocalStorage loaded from %s", path)
