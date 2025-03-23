import vk_api
import time
import re
import os
import requests
from openpyxl import Workbook

VK_URL_PREFIX = "https://vk.com/id"
AVATAR_FOLDER = "avatars"

def get_user_id(vk, screen_name):
    try:
        user = vk.users.get(user_ids=screen_name)
        return user[0]["id"]
    except Exception as e:
        print(f"Ошибка получения ID пользователя: {e}")
        return None

def download_avatar(photo_url, filename):
    try:
        response = requests.get(photo_url)
        with open(filename, "wb") as f:
            f.write(response.content)
    except Exception as e:
        print(f"Не удалось скачать аватар {photo_url}: {e}")

def get_friends_info(vk, user_id):
    try:
        friends = vk.friends.get(user_id=user_id)
        friends_info = []

        if not os.path.exists(AVATAR_FOLDER):
            os.makedirs(AVATAR_FOLDER)

        for friend_id in friends["items"]:
            time.sleep(0.3)
            friend = vk.users.get(user_ids=friend_id, fields="bdate,sex,city,photo_200_orig")
            if friend:
                f = friend[0]
                avatar_url = f.get("photo_200_orig", "")
                avatar_filename = f"{AVATAR_FOLDER}/id{f.get('id')}.jpg"

                if avatar_url:
                    download_avatar(avatar_url, avatar_filename)

                info = {
                    "id": f.get("id"),
                    "link": f"{VK_URL_PREFIX}{f.get('id')}",
                    "first_name": f.get("first_name", ""),
                    "last_name": f.get("last_name", ""),
                    "sex": "мужской" if f.get("sex") == 2 else "женский" if f.get("sex") == 1 else "не указан",
                    "bdate": f.get("bdate", ""),
                    "city": f.get("city", {}).get("title", "") if f.get("city") else "",
                    "avatar_path": avatar_filename
                }
                friends_info.append(info)
        return friends_info
    except Exception as e:
        print(f"Ошибка получения информации о друзьях: {e}")
        return []

def save_to_excel(filename, friends_info):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Друзья"
        ws.append(["ID", "Ссылка", "Имя", "Фамилия", "Пол", "Город", "Дата рождения", "Аватар (файл)"])

        for f in friends_info:
            ws.append([
                f["id"],
                f["link"],
                f["first_name"],
                f["last_name"],
                f["sex"],
                f["city"],
                f["bdate"],
                f["avatar_path"]
            ])

        wb.save(filename)
        print(f"Сохранено в Excel: {filename}")
    except Exception as e:
        print(f"Ошибка при сохранении Excel: {e}")

def main():
    access_token = input("Введите access_token: ").strip()
    profile_url = input("Введите ссылку на профиль пользователя VK: ").strip()

    if not access_token or not profile_url:
        print("Ошибка: не введён токен или ссылка.")
        return

    session = vk_api.VkApi(token=access_token)
    vk = session.get_api()

    screen_name = profile_url.replace("https://vk.com/", "").replace("vk.com/", "").strip()
    user_id = get_user_id(vk, screen_name)
    if not user_id:
        print("Не удалось получить ID пользователя.")
        return

    print("Получаем основные данные друзей пользователя...")
    friends_info = get_friends_info(vk, user_id)

    if not friends_info:
        print("Нет данных о друзьях.")
        return

    save_to_excel(f"{screen_name}.xlsx", friends_info)

if __name__ == "__main__":
    main()
