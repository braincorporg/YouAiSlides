from flask import Flask, request, send_file
from pptx import Presentation
from pptx.util import Inches
import os
import json

app = Flask(__name__)

@app.route('/create_pptx', methods=['POST'])
def create_pptx():
    data = request.json

    prs = Presentation()

    for slide_data in data.get('slides', []):
        slide_layout = prs.slide_layouts[slide_data['layout']]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        content = slide.placeholders[1]

        title.text = slide_data.get('title', '')

        if 'content' in slide_data:
            content.text = slide_data['content']

        if 'image_path' in slide_data and os.path.exists(slide_data['image_path']):
            left = Inches(2)
            top = Inches(2)
            width = Inches(4)
            height = Inches(3)
            slide.shapes.add_picture(slide_data['image_path'], left, top, width, height)

    pptx_file = 'presentation.pptx'
    prs.save(pptx_file)

    return send_file(pptx_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
