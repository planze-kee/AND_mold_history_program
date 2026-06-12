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
        # 공통 설정 (모든 탭이 공유)
        "db_xlsx":          "data/output/output_from_hwp.xlsx",
        "template":         "data/templates/Word_양식.docx",
        "img_dir":          "img",
        # 탭별 설정
        "hwp_input":        "YES",
        "hwp_output":       "data/output/output_from_hwp.xlsx",
        "img_input":        "YES",
        "docx_output":      "data/output",
        "hist_dir":         "data/output",
        "pdf_batch_input":  "data/output",
        "pdf_batch_output": "data/output_pdf",
        "pdf_merge_output": "data/output_pdf/merged.pdf",
    },
    "ui": {
        "window_x":      100,
        "window_y":      100,
        "window_width":  560,
        "window_height": 640,
    },
}

# 구 버전 config.yaml 경로 키 → 통합 키 매핑 (앞쪽 키 우선)
LEGACY_PATH_MAP: dict = {
    "db_xlsx":  ["docx_xlsx", "hist_xlsx"],
    "template": ["docx_template", "hist_template"],
    "img_dir":  ["docx_img", "hist_img", "img_output"],
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
                loaded = self._migrate_legacy_paths(loaded)
                return self._merge(DEFAULTS, loaded)
            except Exception as e:
                logger.warning(f"config.yaml 로드 실패, 기본값 사용: {e}")
        return copy.deepcopy(DEFAULTS)

    # ------------------------------------------------------------- migration
    @staticmethod
    def _migrate_legacy_paths(loaded: dict) -> dict:
        """구 버전 경로 키(docx_xlsx/hist_xlsx 등)를 통합 키로 1회 변환.

        통합 키가 이미 있으면 그대로 두고, 없으면 LEGACY_PATH_MAP의
        앞쪽 구 키 값을 우선 복사한다. 구 키는 제거한다 (save 시 정리됨).
        """
        paths = loaded.get("paths")
        if not isinstance(paths, dict):
            return loaded
        for new_key, old_keys in LEGACY_PATH_MAP.items():
            if not paths.get(new_key):
                for old in old_keys:
                    if paths.get(old):
                        paths[new_key] = paths[old]
                        break
        for old_keys in LEGACY_PATH_MAP.values():
            for old in old_keys:
                paths.pop(old, None)
        return loaded

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
