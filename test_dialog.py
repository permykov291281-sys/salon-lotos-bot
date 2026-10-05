"""Прогон диалога с Настеной без VK: python3 test_dialog.py"""
import sys; sys.argv.append("--cli")
import importlib.util
spec = importlib.util.spec_from_file_location("bot", "bot.py"); bot = importlib.util.module_from_spec(spec); spec.loader.exec_module(bot)
bot.XLSX = bot.DATA / "test_zapisi.xlsx"; bot.XLSX.unlink(missing_ok=True); bot.history.clear()
for t in sys.argv[1:-1] if len(sys.argv) > 2 else [
    "Привет) хочу маникюр",
    "а сколько стоит?",
    "дороговато конечно",
    "давайте на завтра, после обеда",
    "первый вариант",
    "Лена, 9000000000",
    "да",
]:
    print("\nКлиент:", t); bot.handle(777, t)
