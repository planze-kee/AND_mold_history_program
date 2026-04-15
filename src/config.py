"""
설정 관리 모듈 - config.yaml 로드/저장
"""
import copy
import logging
from pathlib import Path

import yaml

logger = logging.getLogger(__name__)

CONFIG_PATH = Path("config.yaml")

DEFAULTS: dict = {
    "paths": {
        "hwp_input":        "YES",
        "hwp_output":       "data/output/output_from_hwp.xlsx",
        "img_input":        "YES",
        "img_output":       "img",
        "docx_xlsx":        "data/output/output_from_hwp.xlsx",
        "docx_template":    "data/templates/Word_양식.docx",
        "docx_img":         "img",
        "docx_output":      "data/output",
        "hist_xlsx":        "data/output/00.DB_19-000.xlsx",
        "hist_template":    "data/templates/Word_양식.docx",
        "hist_img":         "img",
        "hist_dir":         "data/output",
        "pdf_batch_input":  "data/output",
        "pdf_batch_output": "data/output_pdf",
        "pdf_merge_output": "data/output_pdf/merged.pdf",
    },
    "ui": {
        "window_x":      100,
        "window_y":      100,
        "window_width":  560,
        "window_height": 520,
    },
}


class Config:
    """config.yaml 기반 설정 관리.

    사용 예::

        cfg = Config()
        path = cfg.get("paths", "hwp_input")   # 값 읽기
        cfg.set("paths", "hwp_input", "YES2")  # 값 변경
        cfg.save()                              # 저장
    """

    def __init__(self, path: Path = CONFIG_PATH):
        self._path = path
        self._data = self._load()

    # ------------------------------------------------------------------ load
    def _load(self) -> dict:
        if self._path.exists():
            try:
                with open(self._path, "r", encoding="utf-8") as f:
                    loaded = yaml.safe_load(f) or {}
                return self._merge(DEFAULTS, loaded)
            except Exception as e:
                logger.warning(f"config.yaml 로드 실패, 기본값 사용: {e}")
        return copy.deepcopy(DEFAULTS)

    # ------------------------------------------------------------------ merge
    @staticmethod
    def _merge(defaults: dict, overrides: dict) -> dict:
        """기본값 딕셔너리에 overrides를 재귀적으로 덮어씁니다."""
        result = copy.deepcopy(defaults)
        for key, value in overrides.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = Config._merge(result[key], value)
            else:
                result[key] = value
        return result

    # ------------------------------------------------------------------ API
    def get(self, section: str, key: str, fallback: str = "") -> str:
        """설정 값 반환. 없으면 fallback 반환."""
        return str(self._data.get(section, {}).get(key, fallback))

    def get_int(self, section: str, key: str, fallback: int = 0) -> int:
        """정수 설정 값 반환."""
        try:
            return int(self._data.get(section, {}).get(key, fallback))
        except (TypeError, ValueError):
            return fallback

    def set(self, section: str, key: str, value) -> None:
        """설정 값 변경 (메모리만, save() 호출 전까지 파일 미반영)."""
        if section not in self._data:
            self._data[section] = {}
        self._data[section][key] = value

    def save(self) -> None:
        """현재 설정을 config.yaml에 저장."""
        try:
            with open(self._path, "w", encoding="utf-8") as f:
                yaml.dump(
                    self._data, f,
                    allow_unicode=True,
                    default_flow_style=False,
                    sort_keys=False,
                )
        except Exception as e:
            logger.warning(f"config.yaml 저장 실패: {e}")
