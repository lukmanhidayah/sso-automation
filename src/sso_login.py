from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils import save_cookies, load_cookies, save_local_storage, load_local_storage
import logging
import time
import os
from dotenv import load_dotenv
import json as _json
from urllib.request import urlopen

load_dotenv('config/credentials.env')

def login_sso(driver, config):
    driver.get(config['sso_url'])
    wait = WebDriverWait(driver, config['timeout'])
    
    # Coba load cookies & localStorage untuk skip login
    load_cookies(driver, 'data/sso_cookies.json')
    load_local_storage(driver, 'data/sso_localstorage.json')
    driver.get(config['sso_url'])  # Refresh untuk apply cookies & storage
    
    # Cek apakah sudah login (berdasarkan redirect URL)
    try:
        if wait.until(EC.url_contains(config['redirect_url'])):
            logging.info("Already logged in via cookies.")
            return True
    except:
        pass
    
    # Isi form login
    try:
        username_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, config['username_field'])))
        password_field = driver.find_element(By.CSS_SELECTOR, config['password_field'])
        login_button = driver.find_element(By.CSS_SELECTOR, config['login_button'])
        
        username_field.send_keys(os.getenv('SSO_USERNAME'))
        password_field.send_keys(os.getenv('SSO_PASSWORD'))
        login_button.click()
        
        # Setelah submit kredensial, tunggu salah satu kondisi:
        # - Redirect ke redirect_url (berarti tanpa OTP)
        # - Muncul field OTP (berarti perlu OTP)
        try:
            wait.until(EC.any_of(
                EC.url_contains(config['redirect_url']),
                EC.presence_of_element_located((By.CSS_SELECTOR, '#otp'))
            ))
        except Exception:
            # Jika tidak ada salah satunya dalam timeout, anggap gagal
            raise

        # Jika halaman meminta OTP
        if driver.current_url and 'login-actions/authenticate' in driver.current_url:
            try:
                otp_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#otp')))
            except Exception:
                # Fallback: cari by name
                otp_input = wait.until(EC.presence_of_element_located((By.NAME, 'otp')))

            # Ambil account info dulu untuk memilih device/credential
            totp_account = None
            # Allow override via env var; fallback to config, then sensible default for Docker
            totp_url = os.getenv('TOTP_URL') or config.get('totp_url', 'http://host.docker.internal:8001/totp')
            try:
                with urlopen(totp_url, timeout=5) as resp:
                    data = resp.read().decode('utf-8')
                    totp_payload = _json.loads(data)
                    totp_account = str("wasis kurniawan").strip()
            except Exception as ex:
                logging.error("Failed fetching TOTP for account info: %s", str(ex))

            # Jika ada pilihan device/credential, coba pilih sesuai account
            try:
                if totp_account:
                    titles = driver.find_elements(By.CSS_SELECTOR, '.pf-c-tile__title')
                    for t in titles:
                        try:
                            if t.text and t.text.strip() == totp_account:
                                t.click()
                                time.sleep(2)  # Tambah delay agar halaman sempat termuat
                                break
                        except Exception:
                            continue
            except Exception:
                pass

            # Retry submit OTP jika gagal redirect
            for attempt in range(3):
                try:
                    # Ambil TOTP dari service lokal (fresh setiap attempt)
                    with urlopen(totp_url, timeout=5) as resp:
                        data = resp.read().decode('utf-8')
                        totp_payload = _json.loads(data)
                        totp_value = totp_payload.get('totp')
                    if not totp_value:
                        raise RuntimeError('No TOTP value received from provider')

                    otp_input.clear()
                    logging.info(f"Filling OTP input with value (attempt {attempt+1}): {str(totp_value)}")
                    otp_input.send_keys(str(totp_value))
                    time.sleep(2)  # Tambah delay sebelum submit OTP

                    submit_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#kc-login')))
                    submit_btn.click()

                    # Tunggu redirect setelah OTP
                    wait.until(EC.url_contains(config['redirect_url']))
                    break  # Berhasil, keluar dari loop
                except Exception as e:
                    logging.warning(f"OTP submit failed (attempt {attempt+1}): {str(e)}")
                    if attempt == 2:
                        raise  # Sudah 3x gagal, raise error
                    time.sleep(1)  # Tunggu sebentar sebelum retry
        else:
            # Tidak perlu OTP; sudah redirect
            pass

        logging.info("Login successful, redirected to %s", driver.current_url)
        
        # Simpan cookies & localStorage untuk sesi berikutnya
        save_cookies(driver, 'data/sso_cookies.json')
        save_local_storage(driver, 'data/sso_localstorage.json')
        return True
    except Exception as e:
        logging.error("Login failed: %s", str(e))
        driver.save_screenshot('data/logs/error.png')  # Debug
        return False
