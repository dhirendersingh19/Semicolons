from flask import Flask, jsonify, request, Blueprint
from pymongo import MongoClient
import bcrypt
from flask_jwt_extended import JWTManager, create_access_token, jwt_required, get_jwt_identity
from bson.json_util import dumps

meeting_routes = Blueprint('meeting_routes', __name__)

# MongoDB connection string and database name
uri = 'localhost:27017'
db_name = 'Semicolons'

#Storing meeting notes
@meeting_routes.route('/meeting-notes', methods=['POST'])
def create_meeting_notes():
    try:

        # Connect to MongoDB
        client = MongoClient(uri)
        db = client[db_name]

        data = request.get_json()
        username = data['username']
        file_name = data['file_name']
        transcript = data['transcript']
        notes = data['notes']

        #inser meeting notes into database
        existing = db.meeting.find_one({'username':username,'file_name':file_name})
        if existing:
            return jsonify({'message':'File name already exists'})
        else:
            db.meeting.insert_one({
            'username': username,
            'file_name': file_name,
            'transcript' :  transcript,
            'notes': notes,
            'summary': "",
            'completed_task': [],
            'in_progress_task': [],
            'test_cases': [],
            'sprint_work': []
            })

            response = jsonify({'message':'notes added successfully'})
            response.status_code = 201
            return response

    except Exception as e:
        print('Failed to create user:', e)
        response = jsonify({'error': 'Failed to create user'})
        response.status_code = 500
        return response
        
    finally:
        # Close the MongoDB client
        client.close()

#update test case array
@meeting_routes.route('/update-test-cases', methods=['PUT'])
def update_test_case():
    try:
        # Connect to MongoDB
        client = MongoClient(uri)
        db = client[db_name]

        data = request.get_json()
        username = data['username']
        file_name = data['file_name']
        test_case_name = data['test_case_name']
        test_case_status = data['test_case_status']
        test_executed_by = data['test_executed_by']

        result = db.meeting.update_one(
            {'username':username, 'file_name':file_name},
            {'$push': {'test_cases':{
                'name': test_case_name,
                'status': test_case_status,
                'executed_by': test_executed_by
            }}}
        )
        # Send a response indicating success
        response = jsonify({'message': 'test case updated successfully'})
        response.status_code = 201
        return response
            
    except Exception as e:
        print('Failed to update test case:', e)
        response = jsonify({'error': 'Failed to update test case'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()

#update summary 
@meeting_routes.route('/update-summary', methods=['PUT'])
def update_summary():
    try:
        # Connect to MongoDB
        client = MongoClient(uri)
        db = client[db_name]
        data = request.get_json()
        username = data['username']
        file_name = data['file_name']
        summary = data['summary']

        result = db.meeting.update_one(
            {'username':username, 'file_name':file_name},
            {'$set' :{'summary':summary}}
        )

        # Send a response indicating success
        response = jsonify({'message': 'summary updated successfully'})
        response.status_code = 201
        return response

    except Exception as e:
        print('Failed to update summary:', e)
        response = jsonify({'error': 'Failed to update summary'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()  

#update completed task
@meeting_routes.route('/update-completed-task', methods=['PUT'])
def update_completed_task():
    try:
        # Connect to MongoDB
        client = MongoClient(uri)
        db = client[db_name]

        data = request.get_json()
        username = data['username']
        file_name = data['file_name']
        completed_task = data['completed_task']

        for complete in completed_task:
            result = db.meeting.update_one(
                {'username':username, 'file_name':file_name},
                {'$push' :{'completed_task':complete}}
            )

        # Send a response indicating success
        response = jsonify({'message': 'completed task updated successfully'})
        response.status_code = 201
        return response

    except Exception as e:
        print('Failed to update summary:', e)
        response = jsonify({'error': 'Failed to update completed task'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()   

#update pending task
@meeting_routes.route('/update-pending-task', methods=['PUT'])
def update_pending_task():
    try:
        # Connect to MongoDB
        client = MongoClient(uri)
        db = client[db_name]

        data = request.get_json()
        username = data['username']
        file_name = data['file_name']
        pending_task = data['pending_task']

        for pending in pending_task:
            result = db.meeting.update_one(
                {'username':username, 'file_name':file_name},
                {'$push' :{'in_progress_task':pending}}
            )

        # Send a response indicating success
        response = jsonify({'message': 'pending task updated successfully'})
        response.status_code = 201
        return response

    except Exception as e:
        print('Failed to update summary:', e)
        response = jsonify({'error': 'Failed to update pending task'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()

#update sprint work
@meeting_routes.route('/update-sprint-work', methods=['PUT'])
def update_sprint_work():
    try:
        # Connect to MongoDB
        client = MongoClient(uri)
        db = client[db_name]

        data = request.get_json()
        username = data['username']
        file_name = data['file_name']
        sprint_executed_by = data['sprint_executed_by']
        sprint_completed_task = data['sprint_completed_task']  

        result = db.meeting.update_one(
            {'username':username, 'file_name':file_name},
            {'$push': {'sprint_work':{
                'name': sprint_executed_by,
                'completed_task': sprint_completed_task
            }}}
        )
        # Send a response indicating success
        response = jsonify({'message': 'test case updated successfully'})
        response.status_code = 201
        return response

    except Exception as e:
        print('Failed to update sprint work:', e)
        response = jsonify({'error': 'Failed to update sprint work'})
        response.status_code = 500
        return response

    finally:
        # Close the MongoDB client
        client.close()