import argparse
import asyncio
import discord
import json
import os
import schedule
import threading
import time
import xlsxwriter
from datetime import datetime, timedelta

intents = discord.Intents.default()
intents.members = True
client = discord.Client(intents=intents)
config_file = "" # Changed on runtime by calling the bot with the '-c' and a config file path
user_data = [] # Grabbed initially on initilization of bot, this stores all the data of the discord members, is later appended to whenever someone joins
whitelist = [] # Loaded from config
bots = [] # Loaded from config, also appended to by the bot itself
TOKEN = "" # Loaded from config
HIDDEN_CHANNEL_ID = "" # Loaded from config
GLOW_THRESHOLD = 30 # Loaded from config

# Formats a datetime object to something like: 2-31-2023
def str_time(time_object):
    if not type(time_object) == None:
        return time_object.strftime("%Y-%m-%d")

# Calculates days since last post by subtracting the current date and time from the DiscordMember.last_post
def days_since_post(last_post):
    if not last_post == None:
        return (datetime.now().astimezone() - last_post).days + 1
    else:
        return "No posts recorded"

# Generates the name for the xlsx sheet to be generated and uploaded, uses datetime and str_time
def generate_file_name():
    return f'discord_datasheet_{str_time(datetime.now())}.xlsx'


# It's literally in the name, is the DiscordMember class. Handles elements of each discord user
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

        # Glow Points should be thought of as a suspicion level towards a user. Golf rules, higher is bad.
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

    # Outputs the values in a semi-readable and usable format (similar to JSON), is later called in the process of generating the xlsx sheet
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

# Reads in the config file, and initializes a few variables with values from the config file
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

# Using argparse and the '-c' flag, this function is used to initialize the config file location 
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

# Grabs all current users, grabs relevant data, throws it into the list user_data
async def first_run():
    for guild in client.guilds:
        for member in guild.members:
            user_data.append(DiscordMember(member.id, member.name, member.joined_at, member.created_at))
    await run_daily()

# Creates an xlsx file from data within user_data
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

# Uploads xlsx sheet to the channel specified within the config file
async def upload_excel_sheet(channel_id):
    await client.wait_until_ready()
    channel = client.get_channel(channel_id)
    await channel.send(file=discord.File(generate_file_name()))
    await asyncio.sleep(10)

# Initializes the scheduled task, which is running the run_daily function
def timed_functionality():
    schedule.every().day.at("20:00").do(run_daily)
    while True:
        schedule.run_pending()
        time.sleep(1)

# Ran daily by timed_functionality, this generates an xlsx sheet and then uploads it to the channel specified within the config
async def run_daily():
    generate_excel_sheet(user_data)
    await upload_excel_sheet(HIDDEN_CHANNEL_ID)
    os.remove(generate_file_name())
    
# Handles initial connection of the bot, adds itself to the bots list, then calls the first_run function
@client.event
async def on_ready():
    print(f'Connected to Discord with user {client.user}:{client.user.id}')
    bots.append(client.user.id)
    await first_run()
    for member in user_data:
        print(member.__enumerate__())

# Handles new members joining, adds new users to user_data, then messages the channel specified in the config with username and glow_score
@client.event
async def on_member_join(member):
    user_data.append(DiscordMember(member.id, member.name, member.joined_at, member.created_at))
    await client.wait_until_ready()
    channel = client.get_channel(HIDDEN_CHANNEL_ID)
    new_member = user_data[-1]
    await channel.send(f"New member: {new_member['name']}\nGlow Score: {new_member['glow_score']}")

# Handles new messages, adds +1 to a member object when they send a message, also updates their last_post to reflect the date of the message
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

# Starts thread for the timed_functionality function, runs the discord client (starting the bot)
def main(token):
    thread = threading.Thread(target=timed_functionality)
    thread.start()
    client.run(token)

# Entry point. Grabs the config location from argument parser. Reads config file and loads the values, calls main
if __name__ == '__main__':
    config_location = argument_parser_init()
    read_config(config_location)
    main(TOKEN)
