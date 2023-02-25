from flask import Flask  
from Routes.slides import app_routes
from Routes.chatgpt import chatgpt_routes
from Routes.user import mongo_routes
from Routes.dalle import dalle_routes
from Routes.meeting import meeting_routes
from flask_jwt_extended import JWTManager, create_access_token, jwt_required, get_jwt_identity
from flask_cors import CORS
  
app = Flask(__name__) #creating the Flask class object
CORS(app)   
app.register_blueprint(app_routes)
app.register_blueprint(chatgpt_routes)
app.register_blueprint(mongo_routes)
app.register_blueprint(dalle_routes)
app.register_blueprint(meeting_routes)

# Configure JWT settings
app.config['JWT_SECRET_KEY'] = 'mysecretkey' # Replace with a secret key of your choice
jwt = JWTManager(app)

if __name__ =='__main__':  
    app.run(debug = True)  