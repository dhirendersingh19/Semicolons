from flask import Flask, jsonify, request, Blueprint
from pymongo import MongoClient
import bcrypt
from flask_jwt_extended import JWTManager, create_access_token, jwt_required, get_jwt_identity
from bson.json_util import dumps

mongo_routes = Blueprint('mongo_routes', __name__)

# MongoDB connection string and database name
uri = 'localhost:27017'
db_name = 'Semicolons'

# Create a new user
@mongo_routes.route('/users', methods=['POST'])
def create_user():
    try:
        # Extract user data from request body
        data = request.get_json()
        username = data['username']
        firstname = data['firstname']
        lastname = data['lastname']
        email = data['email']
        password = data['password']
        rpassword = data['rpassword']

        # Encrypt the password using bcrypt
        hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())

        # Connect to MongoDB
        client = MongoClient(uri)
        db = client[db_name]

        user = db.users.find_one({'username': username})
        if user:
            return jsonify({"message":"Username already exists"})

        # Insert the new user into the "users" collection
        if (password == rpassword): #check to compare password and confirm password.
            db.users.insert_one({
            'username': username,
            'firstname': firstname,
            'lastname': lastname,
            'email': email,
            'password': hashed_password
            })

            # Send a response indicating success
            response = jsonify({'message': 'User created successfully'})
            response.status_code = 201
            return response
        else:
            return jsonify({'message': 'password and confirm password do not match'})

    except Exception as e:
        print('Failed to create user:', e)
        response = jsonify({'error': 'Failed to create user'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()

@mongo_routes.route('/user/<username>', methods=['GET'])
@jwt_required()
def get_user(username):
    try:
        # Get the current user identity from the JWT token
        current_user = get_jwt_identity()

        # Connect to MongoDB and retrieve user data
        client = MongoClient(uri)
        db = client[db_name]
        user = db.users.find_one({'username': username})

        # If the user is found, return their data as a JSON response
        if user:
            response = dumps(user)
            return response

        # If the user is not found, return a 404 error
        response = jsonify({'error': 'User not found'})
        response.status_code = 404
        return response

    except Exception as e:
        print('Failed to retrieve user:', e)
        response = jsonify({'error': 'Failed to retrieve user'})
        response.status_code = 500
        return response

@mongo_routes.route('/users', methods=['GET'])
@jwt_required()
def get_users():
    try:
        # Get the current user identity from the JWT token
        current_user = get_jwt_identity()

        # Connect to MongoDB and retrieve all user data
        client = MongoClient(uri)
        db = client[db_name]
        users = db.users.find()

        # If any users are found, return their data as a JSON response
        if users:
            response = dumps(users)
            return response

        # If no users are found, return a 404 error
        response = jsonify({'error': 'No users found'})
        response.status_code = 404
        return response

    except Exception as e:
        print('Failed to retrieve users:', e)
        response = jsonify({'error': 'Failed to retrieve users'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()

# Delete user by username
@mongo_routes.route('/users/<username>', methods=['DELETE'])
@jwt_required()
def delete_user(username):
    try:
        # Connect to MongoDB and delete the user with the specified username
        client = MongoClient(uri)
        db = client[db_name]
        result = db.users.delete_one({'username': username})

        # If the deletion was successful, return a success message
        if result.deleted_count == 1:
            response = jsonify({'message': 'User deleted successfully'})
            response.status_code = 200
            return response

        # If the user was not found, return a 404 error
        response = jsonify({'error': 'User not found'})
        response.status_code = 404
        return response

    except errors.PyMongoError as e:
        print('Failed to delete user:', e)
        response = jsonify({'error': 'Failed to delete user'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()

# Authenticate a user
@mongo_routes.route('/login', methods=['POST'])
def login():
    try:

        # Connect to MongoDB and retrieve user data
        client = MongoClient(uri)
        db = client[db_name]
        
        # Extract user data from request body
        data = request.get_json()
        username = data['username']
        password = data['password']

        user = db.users.find_one({'username': username})

        # Check if the user exists and the password is correct
        if user and bcrypt.checkpw(password.encode('utf-8'), user['password']):
            # Generate a JWT token and send it as a response
            access_token = create_access_token(identity=username)
            response = jsonify({'access_token': access_token})
            response.status_code = 200
            response.headers.add("Access-Control-Allow-Origin", "*")
            return response

        # Send a response indicating authentication failure
        response = jsonify({'error': 'Invalid username or password'})
        response.status_code = 401
        response.headers.add("Access-Control-Allow-Origin", "*")
        return response

    except Exception as e:
        print('Failed to authenticate user:', e)
        response = jsonify({'error': 'Failed to authenticate user'})
        response.status_code = 500
        response.headers.add("Access-Control-Allow-Origin", "*")
        return response

    finally:
        # Close the MongoDB client
        client.close()
