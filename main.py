import discord
import yaml
from discord.ext import commands
from typing import Union
import sqlite3
from io import BytesIO
import openpyxl
conn = sqlite3.connect('database.db', timeout=5.0)
c = conn.cursor()
conn.row_factory = sqlite3.Row

c.execute('''CREATE TABLE IF NOT EXISTS words (`word` TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS winners (`member_id` INT, `word` TEXT)''')


with open("config.yml", "r") as stream:
    yaml_data = yaml.safe_load(stream)

client = commands.Bot(command_prefix="!", help_command=None)

def get_winner_data() -> openpyxl.Workbook:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet["A1"] = "Discord ID"
    sheet["B1"] = "Word"
    c.execute("SELECT * FROM winners")
    for i,(discord_id,word) in enumerate(c.fetchall()):
        sheet[f"A{i + 2}"] = str(discord_id)
        sheet[f"B{i + 2}"] = word
    return workbook

def clear_winner_data() -> None:
    c.execute("DELETE FROM winners")
    conn.commit()

def get_current_words() -> list[str]:
    c.execute("SELECT * FROM words")
    return [i[0] for i in c.fetchall()]

def add_words(words: list[str]) -> None:
    for word in words:
        c.execute(f"INSERT INTO words VALUES (?)",(word,))
    conn.commit()

def count_correct_guess(word: str, member: discord.Member) -> None:
    c.execute("DELETE FROM words WHERE word = ?",(word,))
    c.execute("INSERT INTO winners VALUES (?,?)",(member.id,word))
    conn.commit()

def look_for_words(message: discord.Message, words: list[str]) -> Union[str,None]:
    for word in words:
        if word in message.content.split(" "):
            return word

async def send_message(ctx: Union[discord.TextChannel,discord.User,discord.Member,commands.Context], data: Union[str,dict], file: openpyxl.Workbook = None, filename: str = "data.xlsx",**vars_for_data) -> discord.Message:
    try:
        newdata = data.copy()
    except AttributeError:
        newdata = str(data)
    if file:
        with BytesIO() as f:
            file.save(f)
            f.seek(0)
            file = discord.File(fp=f, filename=filename)
    if isinstance(newdata,str):
        return await ctx.send(newdata.format(**vars_for_data),file=file)
    else:
        if "description" in newdata.keys():
            newdata["description"] = newdata["description"].format(**vars_for_data)
        if "title" in newdata.keys():
            newdata["title"] = newdata["title"].format(**vars_for_data)
        embed = discord.Embed(**newdata)
        if "fields" in newdata.keys():
            for field in newdata["fields"]:
                embed.add_field(**field)
        return await ctx.send(embed=embed,file=file)

@client.event
async def on_ready():
    print("Bot Started!")

@client.event
async def on_message(message):
    if not message.author.bot:
        if message.channel.id == yaml_data["guess_channel_id"]:
            words = get_current_words()
            word = look_for_words(message,words)
            if word:
                count_correct_guess(word,message.author)
                await send_message(message.channel,yaml_data["messages"]["correct_guess"],user=message.author)
        await client.process_commands(message)

@client.command()
@commands.has_permissions(administrator=True)
async def addwords(ctx,*words):
    add_words([str(word) for word in words])
    await send_message(ctx,yaml_data["messages"]["words_added"])

@client.command()
@commands.has_permissions(administrator=True)
async def winnerdata(ctx):
    await send_message(ctx,yaml_data["messages"]["data"],file=get_winner_data(),filename="winners.xlsx")

@client.command()
@commands.has_permissions(administrator=True)
async def clearwinnerdata(ctx):
    clear_winner_data()
    await send_message(ctx,yaml_data["messages"]["clear_data"])

client.run(yaml_data["Token"])
