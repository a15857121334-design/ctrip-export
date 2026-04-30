from __future__ import annotations

import argparse
import sys

from main import ConfigError
from main import import_playwright
from main import is_placeholder_url
from main import load_config
from main import nested
from main import require_confirmation
from main import wait_for_page_idle


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="手动登录携程后台并保存 Playwright 登录态。")
    parser.add_argument("--config", default="config.yaml", help="配置文件路径。")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        config = load_config(args.config)
        url = config.raw.get("order_page_url")
        if is_placeholder_url(url):
            raise ConfigError("请先在 config.yaml 中填写真实的 order_page_url。")

        exists_text = "会覆盖已有登录态文件" if config.storage_state_path.exists() else "将创建新文件"
        require_confirmation(
            "即将打开浏览器让你手动登录携程后台。\n"
            f"- 登录页面：{url}\n"
            f"- 保存目标：{config.storage_state_path}\n"
            f"- 状态：{exists_text}\n"
            "脚本不会读取或保存账号密码，只保存浏览器登录态。"
        )

        sync_playwright, _ = import_playwright()
        timeout_ms = int(nested(config.raw, "browser", "timeout_ms", default=30000))
        slow_mo_ms = int(nested(config.raw, "browser", "slow_mo_ms", default=0))

        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=False, slow_mo=slow_mo_ms)
            context = browser.new_context()
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            wait_for_page_idle(page, config)

            print("\n请在打开的浏览器中手动完成登录，并确认已经进入订单后台或保持有效登录状态。")
            require_confirmation(
                f"准备保存登录态到：{config.storage_state_path}\n"
                "确认页面已经登录成功后继续。"
            )
            config.storage_state_path.parent.mkdir(parents=True, exist_ok=True)
            context.storage_state(path=str(config.storage_state_path))
            context.close()
            browser.close()

        print(f"已保存登录态：{config.storage_state_path}")
        return 0
    except ConfigError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
