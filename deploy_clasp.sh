#!/bin/bash

# Read each line from the .env file
while IFS='=' read -r key scriptId; do
    # Skip empty lines or lines without a valid scriptId
    if [ -z "$scriptId" ]; then
        continue
    fi

    echo "Processing script ID: $scriptId"
    
    # Set the script ID and push
    clasp setting scriptId "$scriptId"
    clasp push
done < .env
