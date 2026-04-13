#!/usr/bin/env python3
"""
Word 파일 전부 재생성 (이미지 파일명 XLSX 업데이트 포함)
"""
from pathlib import Path
from src.core import DocumentFiller

if __name__ == "__main__":
    xlsx_path = Path("./data/output/00.DB_19-000.xlsx")
    template_path = Path("./data/templates/Word_양식.docx")
    output_dir = Path("./data/output")
    img_dir = Path("./img")

    print("=" * 60)
    print("Word 파일 전부 재생성 시작 (이미지 XLSX 업데이트)")
    print("=" * 60)
    print(f"XLSX: {xlsx_path}")
    print(f"Template: {template_path}")
    print(f"Output: {output_dir}")
    print(f"Images: {img_dir}")
    print()

    # 기존 파일 정리 (필요시)
    # existing = list(output_dir.glob("19-*.docx"))
    # if existing:
    #     print(f"기존 {len(existing)}개 파일 삭제 중...")
    #     import os
    #     from time import sleep
    #     for f in existing:
    #         try:
    #             os.remove(f)
    #             sleep(0.1)
    #         except:
    #             pass

    # Word 파일 생성
    DocumentFiller.process(xlsx_path, template_path, output_dir, img_dir, limit=0)

    print()
    print("=" * 60)
    print("Word generation complete")
    print("=" * 60)
