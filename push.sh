#!/bin/bash
NAME='ITGeeker.net'
echo "This is \"${NAME}\" shell"
read -p "Please enter git update log: "  commit_value
git add --all . && git commit -m "\"${commit_value}\" use quick push" && git push
