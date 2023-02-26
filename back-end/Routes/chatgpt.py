from flask import Flask, request, jsonify, Blueprint,make_response
from pymongo import MongoClient
from bson.json_util import dumps
import openai
import os
import re

# Set up OpenAI API key
openai.api_key = os.environ['OPENAI_KEY']

chatgpt_routes = Blueprint('chatgpt_routes', __name__)

# MongoDB connection string and database name
uri = 'localhost:27017'
db_name = 'Semicolons'

# Define API endpoint for ChatGPT integration
@chatgpt_routes.route('/api/chat', methods=['POST'])
def chat():
    user_input = request.json['message']
    ans={}
    ans['summary']=summary(user_input)
    ans['completed']=completed(user_input)
    ans['inprogress']=inprogress(user_input)
    ans['test_cases']=test_cases(user_input)
    ans['sprints']=sprints(user_input)
    ans['budget']=budget(user_input)

    #print(ans['summary'])
    res =make_response(jsonify(ans), 200)
    return res

#function for Summary
def summary(user_input):
    # Get user input from POST request
    question = "Summarize the sprint"
    
    # Define the parameters for the GPT-3 API request
    prompt = f"{user_input + question}"
    temperature = 0.2
    max_tokens = 1024

    # Call the OpenAI GPT-3 API
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=temperature,
        max_tokens=max_tokens,
        n=1,
        stop=None,
        frequency_penalty=0,
        presence_penalty=0
    )

    # Extract the AI response from the API response
    summary = response.choices[0].text.strip()

    # Return the AI response to the user
    return summary

#function for completed
def completed(user_input):
    # Get user input from POST request
    question = "name of completed tasks by each team member"
    
    # Define the parameters for the GPT-3 API request
    prompt = f"{user_input + question}"
    temperature = 0.2
    max_tokens = 1024

    # Call the OpenAI GPT-3 API
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=temperature,
        max_tokens=max_tokens,
        n=1,
        stop=None,
        frequency_penalty=0,
        presence_penalty=0
    )

    # Extract the AI response from the API response
    completed = response.choices[0].text.strip()
    
    comp_dict = {}
    for task in re.findall(r'(\w+):\s*(.*)\s*\(JIRA-ticket:(\d+)\)', completed):
        comp_dict[task[0]] = f"{task[1]} (JIRA-ticket:{task[2]})"
    
    return comp_dict
    

#function for inprogress
def inprogress(user_input):
    # Get user input from POST request
    question = "name of tasks in progress by each team member"
    
    # Define the parameters for the GPT-3 API request
    prompt = f"{user_input + question}"
    temperature = 0.2
    max_tokens = 1024

    # Call the OpenAI GPT-3 API
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=temperature,
        max_tokens=max_tokens,
        n=1,
        stop=None,
        frequency_penalty=0,
        presence_penalty=0
    )

    # Extract the AI response from the API response
    inprogress = response.choices[0].text.strip()

    inpro_dict = {}
    for task in re.findall(r'(\w+):\s*(.*)\s*\(JIRA-ticket:(\d+)\)', inprogress):
        inpro_dict[task[0]] = f"{task[1]} (JIRA-ticket:{task[2]})"
    
    return inpro_dict

#function for test_cases
def test_cases(user_input):
    # Get user input from POST request
    question = "what were the test cases,their status in pass or fail and assigned to which team member in the sprint?give the answer in key value pairs"
    
    # Define the parameters for the GPT-3 API request
    prompt = f"{user_input + question}"
    temperature = 0.2
    max_tokens = 1024

    # Call the OpenAI GPT-3 API
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=temperature,
        max_tokens=max_tokens,
        n=1,
        stop=None,
        frequency_penalty=0,
        presence_penalty=0
    )

    # Extract the AI response from the API response
    test_cases = response.choices[0].text.strip()

    tasks_dict = {}
    for task in re.findall(r'(\w+):\s*([\w\s,]+),\s*JIRA-ticket:(\d+),\s*Status:\s*(\w+),\s*Assigned to:\s*(\w+)', test_cases):
        tasks_dict[task[0]] = {
            'test_case_name': task[1],
            'status': task[3],
            'executed_by': task[4]
        }
    return test_cases

#function for sprints
def sprints(user_input):
    # Get user input from POST request
    question = "count of tasks done by each team member throughout the sprint?"
    
    # Define the parameters for the GPT-3 API request
    prompt = f"{user_input + question}"
    temperature = 0.2
    max_tokens = 1024

    # Call the OpenAI GPT-3 API
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=temperature,
        max_tokens=max_tokens,
        n=1,
        stop=None,
        frequency_penalty=0,
        presence_penalty=0
    )

    # Extract the AI response from the API response
    sprints = response.choices[0].text.strip()

    spr={}

    name=re.findall(r'(\w+):', sprints)
    no_tasks=re.findall(r': (\w+)', sprints)
    cnt=0
    for i in name:
        spr[i]=no_tasks[cnt]
        cnt+=1

    return spr

#functions for budget
def budget(user_input):
    # Get user input from POST request
    question = "give month and estimate? key value pair"
    
    # Define the parameters for the GPT-3 API request
    prompt = f"{user_input + question}"
    temperature = 0.2
    max_tokens = 1024

    # Call the OpenAI GPT-3 API
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=temperature,
        max_tokens=max_tokens,
        n=1,
        stop=None,
        frequency_penalty=0,
        presence_penalty=0
    )

    # Extract the AI response from the API response
    budget = response.choices[0].text.strip()

    # Return the AI response to the user
    return budget