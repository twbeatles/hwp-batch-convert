# hwp-batch-convert

OpenClaw skill for batch conversion of 한컴 한글 문서(HWP/HWPX) to PDF and related formats on Windows.

## Included
- `SKILL.md`
- `scripts/hwp_batch_convert.py`
- `references/hwpmate-reuse-notes.md`
- packaged `hwp-batch-convert.skill`

## Upstream basis
- https://github.com/twbeatles/HwpMate

## Quick example
```powershell
python scripts/hwp_batch_convert.py "C:\docs\hwp" --format PDF --output-dir "C:\docs\pdf"
```
