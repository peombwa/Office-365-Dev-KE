# Office-365-Dev-KE
This repository contains demo apps that showcase how to write Office Add-ins and consuming Microsoft Graph resources using Angular.js
The repo contains two projects:
## 1. Find Time
Find time is a web app that uses the find meeting times API from Microsoft Graph to find the appropriate meeting times for all the attendees of a meeting.
The app uses hello.js to authenticate to Azure AD which returns a token that we then use to call the find meeting times API from Graph.
The project is written using Angular.js and Office UI fabric for the styling.
    
## 2. Yandex
This is a simple Outlook Add-in that translates the content of an email from Russian to English. It used the free Yandex API to handle the translation.
The project is written using Angular.js and Office UI fabric for the styling.
