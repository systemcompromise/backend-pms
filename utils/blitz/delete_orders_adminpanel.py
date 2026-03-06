import requests
from bs4 import BeautifulSoup

SESSION_ID = "poo0ya99uz41rjcgga8kl97wc8fue2r2"
BASE_URL = "https://adminpanel.rideblitz.id"

AWB_NUMBERS = [
    "BEMAWB-00001100367",
    "BEMAWB-00001100368",
    "BEMAWB-00001100369",
    "BEMAWB-00001102996",
]

HEADERS = {
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "accept-language": "id-ID,id;q=0.9,en-US;q=0.8",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/145.0.0.0 Safari/537.36",
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "same-origin",
    "upgrade-insecure-requests": "1",
}

def get_fresh_csrf(session, url):
    """Ambil CSRF token fresh dari halaman GET"""
    resp = session.get(url, headers=HEADERS)
    soup = BeautifulSoup(resp.text, "html.parser")
    token = soup.find("input", {"name": "csrfmiddlewaretoken"})
    if token:
        return token["value"]
    # Fallback dari cookie
    return session.cookies.get("csrftoken")

def delete_order(session, awb_number):
    order_id = awb_number.split("-")[-1].lstrip("0")
    url = f"{BASE_URL}/api/order/{order_id}/change/"

    # Ambil CSRF token fresh dari halaman form
    csrf_token = get_fresh_csrf(session, url)
    if not csrf_token:
        print(f"[ERROR] Gagal ambil CSRF token untuk {awb_number}")
        return

    payload = {
        "csrfmiddlewaretoken": csrf_token,
        "order_status": "15",
        "cancel_reason": "delete",  # lowercase sesuai browser
        "package_weight": "1.00",
        "package_width": "0",
        "package_length": "0",
        "package_height": "0",
        "pickup_address": "JL. Pulo Buaran Raya No. 4, Blok III EE - Kav. No.1, Jakarta, 13930",
        "pickup_postal_code": "12345",
        "pickup_lat": "-6.209309700",
        "pickup_long": "106.915178100",
        "dropoff_address": "JL. PONDASI BLOK S KAV. 36 NO. 27B, RT. 009 RW. 017",
        "dropoff_postal_code": "12345",
        "sender_name": "PT MERAPI UTAMA PHARMA - JK1",
        "sender_phone_number": "+62821660",
        "consignee_name": "300000500690937 - ALPRO PONDASI, AP.",
        "consignee_phone_number": "+62821660",
        "business_hub": "59",
        "_continue": "Save",
    }

    headers = {
        **HEADERS,
        "content-type": "application/x-www-form-urlencoded",
        "origin": BASE_URL,
        "referer": url,
    }

    response = session.post(
        url,
        headers=headers,
        data=payload,
        allow_redirects=True,  # Ikuti redirect!
    )

    if response.status_code == 200:
        print(f"[SUCCESS] {awb_number} -> {response.url}")
    else:
        print(f"[FAILED]  {awb_number} -> Status: {response.status_code}")

def main():
    session = requests.Session()
    # Set session cookie dulu
    session.cookies.set("sessionid", SESSION_ID, domain="adminpanel.rideblitz.id")

    for awb in AWB_NUMBERS:
        delete_order(session, awb)

if __name__ == "__main__":
    main()