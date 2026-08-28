"""一次性脚本：关闭 default_template.docx 中 Heading 样式的自动编号。

背景：模板里 abstractNumId=2 通过 <w:pStyle> 反向绑定到 heading 1~5
（styleId=2/3/4/5/6），同时 heading 1/4/5 的样式定义里直接 <w:numId val="1"/>。
两层绑定导致每个 heading 段落都会被自动注入 第%1章 / %1.%2 / %1.%2.%3 ...
形态的多级编号，与 PRD 手写章节号叠加。

本脚本就地修改模板，移除：
  A. word/numbering.xml 中 abstractNumId=2 的 9 个 <w:lvl> 内 <w:pStyle> 绑定
  B. word/styles.xml 中 heading 1/4/5 样式定义里的 <w:numPr>...</w:numPr> 块

幂等：可重复运行；若条目已删则保持不变。
"""
import os
import re
import shutil
import sys
import zipfile

TEMPLATE = os.path.join(
    os.path.dirname(__file__), '..', 'templates', 'default_template.docx'
)
TEMPLATE = os.path.abspath(TEMPLATE)
BACKUP = TEMPLATE + '.bak'


def patch_numbering_xml(xml: str) -> tuple[str, int]:
    """在 abstractNumId=2 区块内删除 <w:pStyle w:val="2|3|4|5|6"/>。"""
    block_re = re.compile(
        r'(<w:abstractNum\b[^>]*w:abstractNumId="2"[^>]*>)(.*?)(</w:abstractNum>)',
        re.DOTALL,
    )
    pstyle_re = re.compile(
        r'<w:pStyle\b[^/]*w:val="(?:2|3|4|5|6)"\s*/>'
    )
    removed = 0

    def repl(m):
        nonlocal removed
        head, body, tail = m.group(1), m.group(2), m.group(3)
        new_body, n = pstyle_re.subn('', body)
        removed += n
        return head + new_body + tail

    new_xml = block_re.sub(repl, xml, count=1)
    return new_xml, removed


def patch_styles_xml(xml: str) -> tuple[str, int]:
    """删除 styleId in {2,5,6} 的样式块里的 <w:numPr>...</w:numPr> 段落。"""
    style_re = re.compile(
        r'(<w:style\b[^>]*w:styleId="(2|5|6)"[^>]*>)(.*?)(</w:style>)',
        re.DOTALL,
    )
    numpr_re = re.compile(r'<w:numPr\b[^>]*>.*?</w:numPr>', re.DOTALL)
    removed = 0

    def repl(m):
        nonlocal removed
        head, _sid, body, tail = m.group(1), m.group(2), m.group(3), m.group(4)
        new_body, n = numpr_re.subn('', body)
        removed += n
        return head + new_body + tail

    new_xml = style_re.sub(repl, xml)
    return new_xml, removed


def main() -> int:
    if not os.path.exists(TEMPLATE):
        print(f"❌ 找不到模板: {TEMPLATE}", file=sys.stderr)
        return 1

    if not os.path.exists(BACKUP):
        shutil.copy2(TEMPLATE, BACKUP)
        print(f"📦 已备份: {BACKUP}")
    else:
        print(f"ℹ️  备份已存在，跳过: {BACKUP}")

    with zipfile.ZipFile(TEMPLATE, 'r') as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    num_xml = files['word/numbering.xml'].decode('utf-8')
    sty_xml = files['word/styles.xml'].decode('utf-8')

    num_xml_new, n_num = patch_numbering_xml(num_xml)
    sty_xml_new, n_sty = patch_styles_xml(sty_xml)

    print(f"  numbering.xml: 移除 {n_num} 个 <w:pStyle> 绑定（abstractNumId=2 内）")
    print(f"  styles.xml   : 移除 {n_sty} 个 <w:numPr> 段（heading 1/4/5 样式）")

    files['word/numbering.xml'] = num_xml_new.encode('utf-8')
    files['word/styles.xml'] = sty_xml_new.encode('utf-8')

    tmp = TEMPLATE + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    os.replace(tmp, TEMPLATE)
    print(f"✅ 模板已就地更新: {TEMPLATE}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
