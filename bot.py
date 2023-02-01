import argparse
import asyncio
import discord
import json
import schedule
import threading
import time
import xlsxwriter
from datetime import datetime, timedelta

intents = discord.Intents.default()
intents.members = True
client = discord.Client(intents=intents)
config_file = ""
user_data = []
whitelist = []
bots = []
TOKEN = ""
HIDDEN_CHANNEL_ID = ""
GLOW_THRESHOLD = 30

def str_time(time_object):
    if not type(time_object) == None:
        return time_object.strftime("%m-%d-%Y")

def days_since_post(last_post):
    if not last_post == None:
        return (datetime.now().astimezone() - last_post).days + 1
    else:
        return "No posts recorded"

def generate_file_name():
    return f'discord_datasheet_{str_time(datetime.now())}.xlsx'


class DiscordMember:
    def __init__(self, id, name, join_date, create_date):
        self.id = id
        self.name = name
        self.join_date = join_date
        self.create_date = create_date
        self.message_count = 0
        self.last_post = None
        if self.id == client.user.id:
            self.bot = True
        elif self.id in bots:
            self.bot = True
        else:
            self.bot = False

        if not self.id in whitelist and not self.id in bots:
            self.whitelisted = False
            if str_time(self.create_date) == str_time(self.join_date):
                self.glow_points = 2
            elif self.create_date + timedelta(days = GLOW_THRESHOLD) >= self.join_date:
                self.glow_points = 1
            else:
                self.glow_points = 0
        else:
            self.whitelisted = True
            self.glow_points = -1

    def __enumerate__(self):
        return {
                'id':str(self.id),
                'name':self.name,
                'join_date':str_time(self.join_date),
                'create_date':str_time(self.create_date), 
                'message_count':self.message_count,
                'days_since_last_post':days_since_post(self.last_post),
                'glow_points':self.glow_points,
                'whitelisted':self.whitelisted,
                'bot':self.bot
                
                }


def read_config(cfg_path):
    global TOKEN, HIDDEN_CHANNEL_ID, GLOW_THRESHOLD
    with open(cfg_path, 'r') as f:
        field = json.load(f)
        TOKEN = field['token']
        GLOW_THRESHOLD = field['glow_threshold']
        HIDDEN_CHANNEL_ID = field['hidden_channel_id']
        for user in field['whitelisted_users']:
            whitelist.append(user)
        for bot in field['bots']:
            bots.append(bot)

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

def first_run():
    for guild in client.guilds:
        for member in guild.members:
            user_data.append(DiscordMember(member.id, member.name, member.joined_at, member.created_at))

def generate_excel_sheet(discord_members):
    file_name = generate_file_name()
    workbook = xlsxwriter.Workbook(str(file_name), options={'remove_timezone': True})
    worksheet = workbook.add_worksheet()
    keys = list(discord_members[0].__enumerate__().keys())
    for key in keys:
        worksheet.write(0,keys.index(key), key)

    for member in discord_members:
        for key in member.__enumerate__():
            worksheet.write(discord_members.index(member)+1, list(member.__enumerate__().keys()).index(key), member.__enumerate__()[key])
    workbook.close()


async def upload_excel_sheet(channel_id):
    await client.wait_until_ready()
    channel = client.get_channel(channel_id)
    await channel.send(file=discord.File(generate_file_name()))
    await asyncio.sleep(10)

def timed_functionality():
    schedule.every().day.at("20:00").do(run_daily)
    while True:
        schedule.run_pending()
        time.sleep(1)

async def run_daily():
    generate_excel_sheet(user_data)
    await upload_excel_sheet(HIDDEN_CHANNEL_ID)
    
@client.event
async def on_ready():
    print(f'Connected to Discord with user {client.user}:{client.user.id}')
    bots.append(client.user.id)
    first_run()
    print(whitelist)
    for member in user_data:
        print(member.__enumerate__())

@client.event
async def on_member_join(member):
    user_data.append(DiscordMember(member.id, member.name, member.joined_at, member.created_at))
    await client.wait_until_ready()
    channel = client.get_channel(HIDDEN_CHANNEL_ID)
    new_member = user_data[-1]
    await channel.send(f"New member: {new_member['name']}\nGlow Score: {new_member['glow_score']}")

@client.event
async def on_message(message):
    channel = message.channel
    if message.author.id == client.user.id:
        return
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