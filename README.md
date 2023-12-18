# Sprint_Automator-Preenchimento_de_Planilhas

## Overview

The Sprint Automator is a project designed to automate the task of populating cells with information related to tasks completed during sprints.

## Objective

The main goal of this project is to streamline and simplify the process of updating and maintaining sprint-related data in spreadsheets. By automating the data entry tasks, teams can save time and reduce the risk of errors associated with manual input.

## Features

- **Efficient Data Entry:** Automate the filling of cells with relevant information.
- **Task Tracking:** Keep track of tasks completed during sprints effortlessly.
- **Time Savings:** Reduce the time and effort required for manual data entry.

## How to Use

1. Clone this repository to your local machine.
   ```bash
   git clone https://github.com/your-username/Sprint_Automator-Preenchimento_de_Planilhas.git
2. You will need a google service account:
- Create one and create a new project.
- Activate Google Sheets API e Google Drive API. 
- Click on APIs and services, credentials, create credentials, service account.
- Fill the blanks and give owner role
- After creating of the new serice account, click actions, manage keys, add key, create new key, JSON and create.

The json file will be downloaded with the creadentials to acess the service account. Put the file on the creadentials folder of this project and rename it to credentials.json

You also will need to give permission to acesse the google spredsheet manually. Go to your spredsheet and click on share. Copy the email id from your service account like: "xxxxxx@yyyyyy.iam.serviceaccount.com" and paste it to send the request to share the spredsheet as an editor.