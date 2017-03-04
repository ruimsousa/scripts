#!/bin/bash

testString="abcdefg:12345:67890:abcde:12345:abcde"
extractedNum="${testString#*:}"     # Remove through first :
echo $extractedNum
extractedNum="${extractedNum#*:}"   # Remove through second :
echo $extractedNum
extractedNum="${extractedNum%%:*}"  # Remove from next : to end of string
echo $extractedNum



testString="DeviceURI socket://prtpt01.lux.eproseed.com:9100"
echo $testString
extractedNum="${testString#*//}"     # Remove through first :
echo $extractedNum
#extractedNum="${extractedNum#*:}"   # Remove through second :
#echo $extractedNum
extractedNum="${extractedNum%%:*}"  # Remove from next : to end of string
echo $extractedNum
