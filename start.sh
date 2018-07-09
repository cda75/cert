#!/bin/bash
SHELL=/bin/bash
export DISPLAY=:99
/etc/init.d/xvfb start
date
echo "Xvfb started at DISPLAY 99"
date
cd /home/dimka/cert
/usr/bin/python /home/dimka/cert/cisco.py >> /home/dimka/cert/cisco.log 
/etc/init.d/xvfb stop
date
echo "Finish checking"

