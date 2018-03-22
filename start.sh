#!/bin/bash

export DISPLAY=:99
/etc/init.d/xvfb start
echo "Xvfb started at DISPLAY 99"
/usr/bin/python cisco.py >> cisco.log  
/etc/init.d/xvfb stop
echo "Finish checking"

