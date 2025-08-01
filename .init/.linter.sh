#!/bin/bash
cd /home/kavia/workspace/code-generation/microsoft-sso-authentication-platform-93221-93231/react_frontend
npm run build
EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ]; then
   exit 1
fi

