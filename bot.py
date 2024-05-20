import discord
from discord.ext import commands
from pymongo import MongoClient
from datetime import datetime, timedelta
import pytz
from dateutil import tz
from bson.objectid import ObjectId
from dotenv import load_dotenv
import os
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Load environment variables from .env file
load_dotenv()

# MongoDB setup
MONGO_URI = os.getenv("MONGO_URI")
client = MongoClient(MONGO_URI)
db = client["working_hours_db"]

# Bot setup
intents = discord.Intents.default()
bot = commands.Bot(command_prefix='/', intents=intents)

# Define timezone from environment variable
TIMEZONE = os.getenv("TIMEZONE")
TIMEZONE_OFFSET = f"Etc/GMT{TIMEZONE.replace(':', '')}"
custom_tz = tz.gettz(TIMEZONE_OFFSET)

def get_collection(guild_id):
    return db[str(guild_id)]

def calculate_total_minutes(entries):
    total_seconds = sum((entry['end_time'] - entry['start_time']).total_seconds() for entry in entries if entry['end_time'])
    return total_seconds / 60

def generate_unique_id(collection, date):
    date_str = date.strftime('%y%m%d')
    count = collection.count_documents({"unique_id": {"$regex": f"^{date_str}-"}}) + 1
    return f"{date_str}-{count:03d}"

def generate_excel(filename, headers, data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    for col_num, header in enumerate(headers, 1):
        cell = sheet.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True)

    for row_num, row_data in enumerate(data, 2):
        for col_num, cell_value in enumerate(row_data, 1):
            sheet.cell(row=row_num, column=col_num, value=cell_value)

    for col_num in range(1, len(headers) + 1):
        sheet.column_dimensions[get_column_letter(col_num)].width = 20

    workbook.save(filename)

@bot.event
async def on_ready():
    print(f'Bot is ready. Logged in as {bot.user}')

@bot.slash_command(name="start", description="Start working")
async def start_work(ctx):

    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    user_id = ctx.author.id
    discord_name = str(ctx.author)
    start_time = datetime.now(tz=custom_tz)
    unique_id = generate_unique_id(collection, start_time)

    # Create a new entry in MongoDB
    entry = {
        "user_id": user_id,
        "discord_name": discord_name,
        "start_time": start_time,
        "end_time": None,
        "unique_id": unique_id
    }
    collection.insert_one(entry)

    embed = discord.Embed(title="Work Start", description=f"Welcome back {ctx.user.mention}!/n Work started at {start_time.strftime('%Y-%m-%d %H:%M')}", color=discord.Color.green())
    embed.add_field(name="Work session ID", value=unique_id)
    embed.set_footer(text="Powered by NickyBoy", icon_url="https://i.imgur.com/QfmDKS6.png")
    await ctx.respond(embed=embed)

@bot.slash_command(name="end", description="End working")
async def end_work(ctx):

    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    user_id = ctx.author.id
    end_time = datetime.now(tz=custom_tz)

    # Find the last entry without an end time
    entry = collection.find_one({"user_id": user_id, "end_time": None})
    if not entry:
        await ctx.respond("No work session to end.")
        return

    collection.update_one({"_id": entry["_id"]}, {"$set": {"end_time": end_time}})

    embed = discord.Embed(title="Work End", description=f"Great work {ctx.author.mention}!\n Work ended at {end_time.strftime('%Y-%m-%d %H:%M')}", color=discord.Color.red())
    embed.set_footer(text="Powered by NickyBoy", icon_url="https://i.imgur.com/QfmDKS6.png")
    await ctx.respond(embed=embed)

@bot.slash_command(name="edit", description="Edit work hours")
async def edit_work(ctx, unique_id: str = None, start_time_new: str = None, end_time_new: str = None):
    await ctx.respond("Alright, let me change the database a little bit.....", ephemeral=True)

    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    if unique_id is None or start_time_new is None or end_time_new is None:
        entries = list(collection.find({"user_id": ctx.author.id}))

        embed = discord.Embed(title="Edit Work Hours", description="Here are your entries. Use /edit [unique_id] [start_time_new] [end_time_new] to edit an entry.", color=discord.Color.blue())
        for entry in entries:
            start = entry['start_time'].astimezone(custom_tz).strftime('%Y-%m-%d %H:%M')
            end = entry['end_time'].astimezone(custom_tz).strftime('%Y-%m-%d %H:%M') if entry['end_time'] else "Ongoing"
            embed.add_field(name=f"ID: {entry['unique_id']}", value=f"Start: {start}\nEnd: {end}", inline=False)
        embed.set_footer(text="Powered by NickyBoy", icon_url="https://i.imgur.com/QfmDKS6.png")
        
        await ctx.respond(embed=embed)
        return

    try:
        new_start_time = datetime.strptime(start_time_new, '%Y-%m-%d %H:%M').replace(tzinfo=custom_tz)
        new_end_time = datetime.strptime(end_time_new, '%Y-%m-%d %H:%M').replace(tzinfo=custom_tz)
    except ValueError:
        await ctx.respond("Invalid date format. Use format: YYYY-MM-DD HH:MM")
        return

    result = collection.update_one({"unique_id": unique_id}, {"$set": {"start_time": start_time_new, "end_time": end_time_new}})
    if result.modified_count == 0:
        await ctx.respond("No entry found with that ID.")
        return

    embed = discord.Embed(title="Work Edited", description=f"Entry {unique_id} has been updated.", color=discord.Color.blue())
    await ctx.respond(embed=embed)

@bot.slash_command(name="check", description="Check your work hours")
async def check_work(ctx):

    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    user_id = ctx.author.id

    entries = list(collection.find({"user_id": user_id}).sort("start_time", -1).limit(10))
    total_minutes = calculate_total_minutes(entries)

    embed = discord.Embed(title="Your Recent Work Hours (Last 10 Entries)", color=discord.Color.gold())
    for entry in entries:
        start = entry['start_time'].astimezone(custom_tz).strftime('%Y-%m-%d %H:%M')
        end = entry['end_time'].astimezone(custom_tz).strftime('%Y-%m-%d %H:%M') if entry['end_time'] else "Ongoing"
        embed.add_field(name=f"ID: {entry['unique_id']}", value=f"Start: {start}\nEnd: {end}", inline=False)
    embed.add_field(name="Total Minutes", value=f"{total_minutes:.2f}", inline=False)
    embed.set_footer(text="Powered by NickyBoy", icon_url="https://i.imgur.com/QfmDKS6.png")
    await ctx.respond(embed=embed, ephemeral=True)

@bot.slash_command(name="list", description="List all work hours")
async def list_work(ctx, start_date: str = None, end_date: str = None):
    await ctx.respond("Ok, let me generate a long list......", ephemeral=True)

    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    try:
        if start_date and end_date:
            start_date = datetime.strptime(start_date, '%Y-%m-%d').replace(tzinfo=custom_tz)
            end_date = datetime.strptime(end_date, '%Y-%m-%d').replace(tzinfo=custom_tz)
            query = {"end_time": {"$gte": start_date, "$lt": end_date + timedelta(days=1)}}
        else:
            query = {}
    except ValueError:
        await ctx.respond("Invalid date format. Use format: YYYY-MM-DD")
        return

    entries = list(collection.find(query))

    user_minutes = {}
    for entry in entries:
        if entry['end_time']:
            user_minutes.setdefault(entry['discord_name'], 0)
            user_minutes[entry['discord_name']] += (entry['end_time'] - entry['start_time']).total_seconds()

    embed = discord.Embed(title="All Users Work Hours", color=discord.Color.purple())
    for user, total_seconds in user_minutes.items():
        total_minutes = total_seconds / 60
        embed.add_field(name=user, value=f"Total Minutes: {total_minutes:.2f}", inline=False)
    embed.set_footer(text="Powered by NickyBoy", icon_url="https://i.imgur.com/QfmDKS6.png")

    await ctx.respond(embed=embed, ephemeral=True)

@bot.slash_command(name="exportdata", description="Export work data to an Excel file")
async def export_data(ctx, start_date: str, end_date: str):
    await ctx.respond("Sure thing! Excel file is coming.....", ephemeral=True)

    guild_id = ctx.guild.id
    collection = get_collection(guild_id)

    try:
        start_date = datetime.strptime(start_date, '%Y-%m-%d').replace(tzinfo=custom_tz)
        end_date = datetime.strptime(end_date, '%Y-%m-%d').replace(tzinfo=custom_tz)
        query = {"end_time": {"$gte": start_date, "$lt": end_date + timedelta(days=1)}}
    except ValueError:
        await ctx.respond("Invalid date format. Use format: YYYY-MM-DD")
        return

    entries = list(collection.find(query))

    data = []
    for entry in entries:
        if entry['end_time']:
            start_time = entry['start_time'].astimezone(custom_tz).strftime('%Y-%m-%d %H:%M')
            end_time = entry['end_time'].astimezone(custom_tz).strftime('%Y-%m-%d %H:%M')
            total_minutes = (entry['end_time'] - entry['start_time']).total_seconds() / 60
            data.append([entry['discord_name'], start_time, end_time, total_minutes])

    headers = ["Employee Name", "Start Time", "End Time", "Total Minutes"]
    filename = f"work_data_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
    generate_excel(filename, headers, data)

    await ctx.respond(file=discord.File(filename))
    os.remove(filename)  # Delete the file after sending

@bot.slash_command(name="exporttotal", description="Export total work hours to an Excel file")
async def export_total(ctx, start_date: str, end_date: str):
    await ctx.respond("Sure thing! Excel file is coming.....", ephemeral=True)

    guild_id = ctx.guild.id
    collection = get_collection(guild_id)

    try:
        start_date = datetime.strptime(start_date, '%Y-%m-%d').replace(tzinfo=custom_tz)
        end_date = datetime.strptime(end_date, '%Y-%m-%d').replace(tzinfo=custom_tz)
        query = {"end_time": {"$gte": start_date, "$lt": end_date + timedelta(days=1)}}
    except ValueError:
        await ctx.respond("Invalid date format. Use format: YYYY-MM-DD")
        return

    entries = list(collection.find(query))

    user_minutes = {}
    for entry in entries:
        if entry['end_time']:
            user_minutes.setdefault(entry['discord_name'], 0)
            user_minutes[entry['discord_name']] += (entry['end_time'] - entry['start_time']).total_seconds()

    data = [[user, total_seconds / 60] for user, total_seconds in user_minutes.items()]
    headers = ["Employee Name", "Total Minutes"]
    filename = f"total_hours_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
    generate_excel(filename, headers, data)

    await ctx.respond(file=discord.File(filename))
    os.remove(filename)  # Delete the file after sending

bot.run(os.getenv("DISCORD_BOT_TOKEN"))
