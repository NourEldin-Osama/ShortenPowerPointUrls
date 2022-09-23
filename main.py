from pptx import Presentation
from pyshorteners import Shortener

FILE_NAME = 'PYPPT.pptx'
prs = Presentation(FILE_NAME)
shortener = Shortener().tinyurl

for slide in prs.slides:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                address = run.hyperlink.address
                if address is None:
                    continue
                # print(run.text)
                try:
                    run.text = shortener.short(address)
                except:
                    continue
                # print(address)
prs.save(FILE_NAME)
