import discord
import xlsxwriter
import argparse
import time
import schedule
import asyncio
from datetime import datetime
import threading
import json

intents = discord.Intents.default()
intents.members = True
client = discord.Client(intents=intents)
user_data = []
whitelist = []
TOKEN = ""
HIDDEN_CHANNEL_ID = ""

def read_config(cfg_path):
    global TOKEN
    with open(cfg_path, 'r') as f:
        field = json.load(f)
        TOKEN = field['token']
        for user in field['whitelisted_users']:
            whitelist.append(user)

def str_time(time_object):
    time_object.strftime("%m-%d-%Y")

def days_since_post(last_post):
    if not last_post == None:
        return (datetime.now() - last_post).days
    else:
        return "No posts recorded"

def first_run():
    for guild in client.guilds:
        for member in guild.members:
            user_data.append(DiscordMember(member.id, member.name, member.joined_at, member.created_at))



def generate_excel_sheet(discord_members):
    return 0
def upload_excel_sheet(filepath, channel):
    return 0

def timed_functionality():
    schedule.every().day.at("20:00").do(run_daily)
    while True:
        schedule.run_pending()
        time.sleep(1)

def run_daily():
    print("function ran successfully")




    
class DiscordMember:
    def __init__(self, id, name, join_date, create_date):
        self.id = id
        self.name = name
        self.join_date = join_date
        self.create_date = create_date
        self.message_count = 0
        self.last_post = None
        if str_time(self.join_date) == str_time(self.create_date) and not self.id in whitelist:
            self.suspicious = True
        else:
            self.suspicious = False

    def __enumerate__(self):
        return {
                'id':self.id,
                'name':self.name,
                'join_date':self.join_date,
                'create_date':self.create_date, 
                'message_count':self.message_count,
                'last_post':self.last_post,
                'suspicious':self.suspicious
                
                }

@client.event
async def on_ready():
    print(f'Connected to Discord with user {client.user}')
    first_run()
    for member in user_data:
        if str_time(member.join_date) == str_time(member.create_date):
            print(f'[Suspicious] Join/Create Day Match: {member.__enumerate__()}')

@client.event
async def on_member_join(member):
    user_data.append(DiscordMember(member.id, member.name, member.joined_at, member.created_at))

@client.event
async def on_message(message):
    channel = message.channel
    for member in user_data:
        if member.id == message.author.id:
            member.message_count += 1
            member.last_post = message.created_at
            print(member.__enumerate__())

def main(token):
    thread = threading.Thread(target=timed_functionality)
    thread.start()
    client.run(token)
    

if __name__ == '__main__':
    read_config("config.json")
    main(TOKEN)