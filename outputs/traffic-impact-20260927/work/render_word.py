import importlib.util
import os
from pathlib import Path
root=Path(__file__).resolve().parent
renderer_path=Path(r'C:\Users\leung\.codex\plugins\cache\openai-primary-runtime\documents\26.905.11957\skills\documents\render_docx.py')
spec=importlib.util.spec_from_file_location('render_docx',renderer_path)
renderer=importlib.util.module_from_spec(spec);spec.loader.exec_module(renderer)
# Use Word's native PDF export because this Windows runtime has no bundled LibreOffice.
pdf=root/'written.pdf'
assert pdf.is_file()
renderer.convert_to_pdf=lambda *args,**kwargs: (str(pdf),'Rendered with Microsoft Word')
poppler=Path(r'C:\Users\leung\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler')
exe=next(poppler.rglob('pdftoppm.exe'))
os.environ['PATH']=str(exe.parent)+os.pathsep+os.environ.get('PATH','')
pages=renderer.rasterize(str(root.parent/'Traffic Impact Written Answers.docx'),str(root/'doc-preview'),150,False,False)
print('\n'.join(pages))
