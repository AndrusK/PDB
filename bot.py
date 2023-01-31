import discord
import argparse
import xlsxwriter
import time
import schedule
from datetime import datetime, timedelta
import threading
import json

intents = discord.Intents.default()
intents.members = True
client = discord.Client(intents=intents)
config_file = ""
user_data = []
whitelist = []
TOKEN = ""
HIDDEN_CHANNEL_ID = ""

class DiscordMember:
    def __init__(self, id, name, join_date, create_date):
        self.id = id
        self.name = name
        self.join_date = join_date
        self.create_date = create_date
        self.message_count = 0
        self.last_post = None

        if not self.id in whitelist:
            if str_time(self.create_date) == str_time(self.join_date):
                self.suspicious_score = 2
            elif self.create_date + timedelta(days = 90) >= self.create_date:
                self.suspicious_score = 1
            else:
                self.suspicious_score = 0
        else:
            self.suspicious_score = -1

    def __enumerate__(self):
        return {
                'id':self.id,
                'name':self.name,
                'join_date':self.join_date,
                'create_date':self.create_date, 
                'message_count':self.message_count,
                'last_post':self.last_post,
                'suspicious_score':self.suspicious_score
                
                }

def read_config(cfg_path):
    global TOKEN
    with open(cfg_path, 'r') as f:
        field = json.load(f)
        TOKEN = field['token']
        for user in field['whitelisted_users']:
            whitelist.append(user)

def argument_parser_init():
    parser = argparse.ArgumentParser(
                    prog = 'PDB (creative name coming soon)',
                    description = 'Simple discord bot to monitor for inactivity and generally suspicous accounts')
    parser.add_argument('-c', '--config', 
                        help='The config file you\'d like to load. Check formatting @ github.com/AndrusK/PDB/config.json',
                        required=True,
                        dest='config_file',
                        action='store'
                        )
    args = vars(parser.parse_args())
    return args["config_file"]

def str_time(time_object):
    return time_object.strftime("%m-%d-%Y")

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

@client.event
async def on_ready():
    print(f'Connected to Discord with user {client.user}')
    first_run()

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
    config_location = argument_parser_init()
    read_config(config_location)
    main(TOKEN)
