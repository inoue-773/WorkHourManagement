import discord
from discord.ext import commands
from pymongo import MongoClient
from datetime import datetime, timedelta
import pytz
from bson.objectid import ObjectId
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# MongoDB setup
MONGO_URI = os.getenv("MONGO_URI")
client = MongoClient(MONGO_URI)
db = client["working_hours_db"]

# Bot setup
intents = discord.Intents.default()
bot = commands.Bot(command_prefix='/', intents=intents)

# Define JST timezone
JST = pytz.timezone('Asia/Tokyo')

def get_collection(guild_id):
    return db[str(guild_id)]

def calculate_total_hours(entries):
    total_seconds = sum((entry['end_time'] - entry['start_time']).total_seconds() for entry in entries if entry['end_time'])
    return total_seconds / 3600

def generate_unique_id(collection, date):
    date_str = date.strftime('%y%m%d')
    count = collection.count_documents({"unique_id": {"$regex": f"^{date_str}-"}}) + 1
    return f"{date_str}-{count:03d}"

@bot.event
async def on_ready():
    print(f'Bot is ready. Logged in as {bot.user}')

@bot.slash_command(name="start", description="Start working")
async def start_work(ctx):
    await ctx.defer()
    
    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    user_id = ctx.author.id
    discord_name = str(ctx.author)
    start_time = datetime.now(JST)
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

    embed = discord.Embed(title="Work Start", description=f"Work started at {start_time.strftime('%Y-%m-%d %H:%M')}", color=discord.Color.green())
    embed.add_field(name="Entry ID", value=unique_id)
    await ctx.send(embed=embed)

@bot.slash_command(name="end", description="End working")
async def end_work(ctx):
    await ctx.defer()
    
    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    user_id = ctx.author.id
    end_time = datetime.now(JST)

    # Find the last entry without an end time
    entry = collection.find_one({"user_id": user_id, "end_time": None})
    if not entry:
        await ctx.send("No work session to end.")
        return

    collection.update_one({"_id": entry["_id"]}, {"$set": {"end_time": end_time}})

    embed = discord.Embed(title="Work End", description=f"Work ended at {end_time.strftime('%Y-%m-%d %H:%M')}", color=discord.Color.red())
    await ctx.send(embed=embed)

@bot.slash_command(name="edit", description="Edit work hours")
async def edit_work(ctx, unique_id: str = None, new_start: str = None, new_end: str = None):
    await ctx.defer()
    
    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    if unique_id is None or new_start is None or new_end is None:
        entries = list(collection.find({"user_id": ctx.author.id}))

        embed = discord.Embed(title="Edit Work Hours", description="Here are your entries. Use /edit [unique_id] [new_start] [new_end] to edit an entry.", color=discord.Color.blue())
        for entry in entries:
            start = entry['start_time'].astimezone(JST).strftime('%Y-%m-%d %H:%M')
            end = entry['end_time'].astimezone(JST).strftime('%Y-%m-%d %H:%M') if entry['end_time'] else "Ongoing"
            embed.add_field(name=f"ID: {entry['unique_id']}", value=f"Start: {start}\nEnd: {end}", inline=False)
        
        await ctx.send(embed=embed)
        return

    try:
        new_start_time = datetime.strptime(new_start, '%Y-%m-%d %H:%M').replace(tzinfo=JST)
        new_end_time = datetime.strptime(new_end, '%Y-%m-%d %H:%M').replace(tzinfo=JST)
    except ValueError:
        await ctx.send("Invalid date format. Use format: YYYY-MM-DD HH:MM")
        return

    result = collection.update_one({"unique_id": unique_id}, {"$set": {"start_time": new_start_time, "end_time": new_end_time}})
    if result.modified_count == 0:
        await ctx.send("No entry found with that ID.")
        return

    embed = discord.Embed(title="Work Edited", description=f"Entry {unique_id} has been updated.", color=discord.Color.blue())
    await ctx.send(embed=embed)

@bot.slash_command(name="check", description="Check your work hours")
async def check_work(ctx):
    await ctx.defer()
    
    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    user_id = ctx.author.id

    entries = list(collection.find({"user_id": user_id}))
    total_hours = calculate_total_hours(entries)

    embed = discord.Embed(title="Your Work Hours", color=discord.Color.gold())
    for entry in entries:
        start = entry['start_time'].astimezone(JST).strftime('%Y-%m-%d %H:%M')
        end = entry['end_time'].astimezone(JST).strftime('%Y-%m-%d %H:%M') if entry['end_time'] else "Ongoing"
        embed.add_field(name=f"ID: {entry['unique_id']}", value=f"Start: {start}\nEnd: {end}", inline=False)
    embed.add_field(name="Total Hours", value=f"{total_hours:.2f}", inline=False)
    await ctx.send(embed=embed)

@bot.slash_command(name="list", description="List all work hours")
async def list_work(ctx, start_date: str = None, end_date: str = None):
    await ctx.defer()
    
    guild_id = ctx.guild.id
    collection = get_collection(guild_id)
    try:
        if start_date and end_date:
            start_date = datetime.strptime(start_date, '%Y-%m-%d').replace(tzinfo=JST)
            end_date = datetime.strptime(end_date, '%Y-%m-%d').replace(tzinfo=JST)
            query = {"end_time": {"$gte": start_date, "$lt": end_date + timedelta(days=1)}}
        else:
            query = {}
    except ValueError:
        await ctx.send("Invalid date format. Use format: YYYY-MM-DD")
        return

    entries = list(collection.find(query))

    user_hours = {}
    for entry in entries:
        if entry['end_time']:
            user_hours.setdefault(entry['discord_name'], 0)
            user_hours[entry['discord_name']] += (entry['end_time'] - entry['start_time']).total_seconds()

    embed = discord.Embed(title="All Users Work Hours", color=discord.Color.purple())
    for user, total_seconds in user_hours.items():
        total_hours = total_seconds / 3600
        embed.add_field(name=user, value=f"Total Hours: {total_hours:.2f}", inline=False)

    await ctx.send(embed=embed)

bot.run(os.getenv("DISCORD_BOT_TOKEN"))
