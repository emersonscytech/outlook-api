#!/bin/bash
DIR_NAME="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )";

$DIR_NAME/node_modules/.bin/tsc -p client;
