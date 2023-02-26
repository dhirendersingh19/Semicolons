from flask import Flask, request, jsonify, Blueprint
from flask_jwt_extended import JWTManager, create_access_token, jwt_required, get_jwt_identity
import openai
import requests
import os

dalle_routes = Blueprint('dalle_routes', __name__)

# Set up OpenAI API key
openai.api_key = os.environ['OPENAI_KEY']

# Define API endpoint for ChatGPT integration
@dalle_routes.route('/api/dalle', methods=['POST'])
@jwt_required()
def chat():
    # Get user input from POST request
    user_input = request.json['description']

    # Call the OpenAI GPT-3 API
    response = openai.Image.create(
    prompt=f"{user_input}",
    model="image-alpha-001",
    size="1024x1024",
    response_format="url"
)

    # Extract the AI response from the API response
    image_url =  response['data'][0]['url']

    # Return the AI response to the user
    response = requests.get(image_url)
    if response.status_code:
        fp = open('download1.png', 'wb')
        fp.write(response.content)
        fp.close()

    return jsonify({'image-url:': image_url})