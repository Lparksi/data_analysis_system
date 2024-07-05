from flask import Flask, request, render_template, send_from_directory
import logging
from loguru import logger
from werkzeug.utils import secure_filename
import os

import data_analysis_system
import time


class InterceptHandler(logging.Handler):
    def emit(self, record):
        level = logger.level(record.levelname).name

        frame, depth = logging.currentframe(), 2
        while frame.f_code.co_filename == logging.__file__:
            frame = frame.f_back
            depth += 1

        # Log the message using Loguru
        logger.opt(depth=depth, exception=record.exc_info).log(level, record.getMessage())


logger.add("DAS.log", rotation="5 MB", enqueue=True)
logging.getLogger('werkzeug').handlers = [InterceptHandler()]
logging.getLogger('werkzeug').propagate = False
app = Flask(__name__)


@app.route('/')
def hello_world():  # put application's code here
    logger.info("HELLO WORD!")
    return render_template('upload.html')


@app.route('/upload', methods=['POST'])
@logger.catch
def handle_upload():
    if 'excel_file' not in request.files:
        return 'No file part'
    if 'config_file' not in request.files:
        config_file = os.path.join('attachment', 'template1.ini')
    else:
        config_file = request.files['config_file']
        if config_file.filename == '':
            config_file = os.path.join('attachment', 'template1.ini')
        else:
            config_file.save(os.path.join('uploads', secure_filename(config_file.filename)))
    file = request.files['excel_file']
    template_option = request.form['template_option']
    if file.filename == '':
        return 'No selected file'
    if file:
        filename = str(time.time()) + secure_filename(file.filename)
        file.save(os.path.join('uploads', filename))
        logger.info("文件上传成功 {}".format(filename))
        ret =  data_analysis_system.compute_excel(os.path.join('uploads', filename), template_option, config_file=config_file)
        return send_from_directory('uploads', ret, as_attachment=True)


@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    return send_from_directory('attachment', filename, as_attachment=True)


if os.path.exists('uploads') is False:
    logger.info("上传目录不存在，创建uploads目录")
    os.mkdir('uploads')
else:
    logger.info("uploads目录已存在")


if __name__ == '__main__':
    app.run()
