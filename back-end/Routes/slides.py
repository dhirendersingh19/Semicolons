import json
from flask import Blueprint, request, make_response, jsonify, send_file
from Utils import utils
from Logs.logging import logger
from Controllers import slides

app_routes = Blueprint('app_routes', __name__)


@app_routes.route(f"/{utils.BASE_ENDPOINT}/v1/slides", methods=['POST'])
def create_slides():
    if request.method == 'POST':

        title = request.json["title"]
        script = request.json["script"]
        logger.debug(f"[DEBUG] Creating slide with title:{title}")

        slide = slides.Slides(utils.BASE_PPT_PATH)
        res, code = slide.create_ppt(title)

        response = make_response(jsonify(res.__dict__), code)
        return response


@app_routes.route(f"/{utils.BASE_ENDPOINT}/v1/slide/<int:id>", methods=['GET'])
def get_slide(id):
    if request.method == 'GET':
        logger.debug("inside get")
        return []


@app_routes.route(f"/{utils.BASE_ENDPOINT}/v1/slides", methods=['GET'])
def get_slides():
    if request.method == 'GET':
        logger.debug("inside get all slides")
        return []


@app_routes.route(f"/{utils.BASE_ENDPOINT}/v1/slides/<int:id>", methods=['DELETE'])
def delete_slide(id):
    if request.method == 'DELETE':
        logger.debug("inside delete")
        return [id]


@app_routes.route(f"/{utils.BASE_ENDPOINT}/v1/download-ppt/<name>")
def download(name):
    logger.debug("inside download")
    path = f"PPT/{name}"
    return send_file(path, as_attachment=True)
