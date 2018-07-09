#!/bin/bash
SHELL=/bin/bash
PATH=/home/dimka/bin:/home/dimka/.local/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/usr/games:/usr/local/games:/snap/bin
export DISPLAY=:99
/etc/init.d/xvfb start
echo "Xvfb started at DISPLAY 99"
date
cd /home/dimka/cert
/usr/bin/python /home/dimka/cert/cisco.py >> /home/dimka/cert/cisco.log 
/etc/init.d/xvfb stop
date
echo "Finish checking"

