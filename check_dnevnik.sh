#!/bin/bash

export DISPLAY=:99
/etc/init.d/xvfb start
echo "Xvfb started at DISPLAY 99"
echo "Start checking diary......"
/usr/bin/python extract.py  
/etc/init.d/xvfb stop
echo "Diary successfully checked"

