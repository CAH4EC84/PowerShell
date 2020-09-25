#!/bin/bash

url="https://example.com/api/v1/chat.postMessage"
channel="$1"
userId="YOUR_USER_ID"
authToken="YOUR_TOKEN"

icon_emoji=':grinning:'
LOGFILE="/var/log/zabbix/zabbix-rocketchat.log"


# Get the Rocketchat zabbix subject ($2 - hopefully either PROBLEM or RECOVERY)
subject="$2"
# Change color emoji depending on the subject - Green (RECOVERY), Red (PROBLEM), Yellow (UPDATE)
if [[ "$subject" == *"OK"* ]]; then
        color="#00ff33"
elif [[ "$subject" == *"UPDATE"* ]]; then
        color="#ffcc00"
elif [[ "$subject" == *"PROBLEM"* ]]; then
        color="#ff2a00"
fi


if [[ "$subject" == *"Resolved"* ]]; then
        icon_emoji=':grinning:'
elif  [[ "$subject" == *"UPDATE"* ]]; then
        icon_emoji=':warning:'
elif  [[ "$subject" == *"Problem: Все работает"* ]]; then
        icon_emoji=':bulb:'
elif  [[ "$subject" == *"Problem"* ]]; then
        icon_emoji=':slight_frown:'

fi


# The message that we want to send to Mattermost is the "subject" value ($2 / $subject - that we got earlier) followed by the message that Zabbix actually sent us ($3)
message="${subject}: $3" 
message="${message//$'\n'/ }"
message="${message//$'\r'/ }"

# Build our JSON payload and send it as a POST request to the Mattermost incoming web-hook URL
payload='{"channel":"'$channel'","emoji":"'$icon_emoji'","attachments":[{"color":"'${color}'","title":"'${subject}'","text":"'${message}'"}]}'
# Send Payload to the Rocket.Chat Server
curl -X POST -H "X-Auth-Token:YOUR_TOKEN" -H "X-User-Id:YOUR_USER_ID" -H "Content-Type: application/json" --data "${payload}"  --insecure $url
echo "$payload" >>$LOGFILE
