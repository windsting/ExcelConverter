#!/bin/bash

pushd `dirname $0` > /dev/null
SCRIPTPATH="$( cd "$(dirname "$0")" ; pwd -P )"
popd > /dev/null

# echo $SCRIPTPATH
DEST_DIR="/usr/bin"
FILE_PATH="$SCRIPTPATH/ExcelConverter"
LINK="$DEST_DIR/ExcelConverter"

chmod +x $FILE_PATH

[ -e $LINK ] && sudo rm $LINK
sudo ln -s $FILE_PATH /usr/bin/
