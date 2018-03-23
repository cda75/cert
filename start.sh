#!/bin/bash
PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/snap/bin
HOME=/home/dimka
export DISPLAY=:99
/etc/init.d/xvfb start
date
echo "Xvfb started at DISPLAY 99"
cd /root/cert
echo $PWD
python cisco.py > /root/cisco.log  
/etc/init.d/xvfb stop
date
echo "Finish checking"

