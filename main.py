from json import load
from time import sleep
from typing import Any, TypedDict

import pyautogui
import pyperclip

# print(pyautogui.position())


def main(profile: Profile) -> None:
    search_btn = pyautogui.locateCenterOnScreen(
        "screens/search_bth.png", confidence=0.9
    )

    move_mouse_and_click(search_btn)
    pyautogui.typewrite(profile["serial_number"], interval=0.3)

    meter_search_btn = pyautogui.locateCenterOnScreen(
        "screens/meter_search_btn.png", confidence=0.9
    )

    move_mouse_and_click(meter_search_btn)
    move_mouse_and_click((783, 497))

    open_btn = pyautogui.locateCenterOnScreen("screens/open_btn.png", confidence=0.9)
    move_mouse_and_click(open_btn)
    # Клик правой кнопкой В Список точек учета по первому ПУ.
    move_mouse_and_click((591, 491), btn="right")
    # Клик по Отобразить данные.
    move_mouse_and_click((640, 390))
    # Клик по кнопке с чёрной стрелкой вниз, где выбор интервала для данных.
    move_mouse_and_click((868, 81))

    a = pyautogui.locateCenterOnScreen("screens/A.png", confidence=0.9)
    move_mouse_and_click(a)
    move_mouse_and_click((1026, 81))

    menu_all = pyautogui.locateCenterOnScreen(
        "screens/all.png", confidence=0.7, grayscale=True
    )

    move_mouse_and_click(menu_all)
    # Дата До
    move_mouse_and_click((521, 82))
    pyautogui.typewrite("1")
    # Дата От
    move_mouse_and_click((451, 85))
    move_mouse_and_click((332, 110))
    # Кнопка с зеленной стрелкой.
    move_mouse_and_click((214, 82), clicks=2)
    # Кнопка с жирной чёрной стрелкой вниз.
    move_mouse_and_click((210, 147))

    create_report = pyautogui.locateCenterOnScreen(
        "screens/create_report.png", confidence=0.9
    )

    move_mouse_and_click(create_report)
    move_mouse_and_click((575, 106))

    excel = pyautogui.locateCenterOnScreen("screens/excel.png", confidence=0.9)
    move_mouse_and_click(excel)
    sleep(2)

    copy_paste_cyrillic(profile["file_name"])
    save_btn = pyautogui.locateCenterOnScreen("screens/save_btn.png", confidence=0.9)
    move_mouse_and_click(save_btn)
    # Кнопка Да для подтверждения для перезаписи файла.
    move_mouse_and_click((1035, 526))
    # Закрыть сохранение файла.
    move_mouse_and_click((465, 50))
    # Закрыть Данные.
    move_mouse_and_click((315, 53))
    sleep(2)


def move_mouse_and_click(coordinates: Any, btn="left", clicks=1) -> None:
    pyautogui.moveTo(coordinates, duration=3)
    pyautogui.click(button=btn, clicks=clicks)


def copy_paste_cyrillic(text: str) -> None:
    pyperclip.copy(text)
    pyautogui.keyDown("ctrl")
    pyautogui.press("v")
    sleep(0.1)
    pyautogui.keyUp("ctrl")


Profile = TypedDict(
    "Profile",
    {
        "file_name": str,
        "serial_number": str,
    },
)

if __name__ == "__main__":
    try:
        with open("profiles.json", mode="r", encoding="UTF-8") as f:
            profiles: list[Profile] = load(f)
    except FileNotFoundError:
        print("Error! File 'profiles.json' not found!")
        exit(1)

    for profile in profiles:
        main(profile)
