# python 3.6+

from dateutil import parser, relativedelta, tz
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
from lh3.api import *

from datetime import datetime
from bs4 import BeautifulSoup
import pendulum
from jinja2 import Environment, FileSystemLoader
import os
import datetime
import logging
from uuid import uuid4
import pandas as pd
from htmldocx import HtmlToDocx
from pprint import pprint as print

new_parser = HtmlToDocx()

logging.basicConfig(format="%(asctime)s %(message)s", datefmt="%m/%d/%Y %I:%M:%S %p")

root = os.path.dirname(os.path.abspath(__file__))
templates_dir = os.path.join(root, "templates")
env = Environment(loader=FileSystemLoader(templates_dir))
template = env.get_template("index.html")


def write_html_to_template(output, filePath):
    """create an HTML file using the default HTML template

    Args:
        output ([string]):  HTML content
    """
    # if file exist

    print("write_html_to_template")

    if os.path.exists(filePath):
        os.remove(filePath)
        with open(filePath, "w", encoding="utf-8") as file:
            file.write(str(output))
    else:
        with open(filePath, "w", encoding="utf-8") as file:
            file.write(str(output))


def retrieve_transcript(transcript_metadata, chat_id):
    """Return a Transcript (dict) containing metadata
        The 'message' is the raw Transcript

    Args:
        transcript_metadata (dict): the Chat Transcript from LibraryH3lp
        chat_id (int): The chat ID

    Returns:
        dict: Return a Transcript (dict) containing metadata
    """
    print("retrieve_transcript")
    queue_id = transcript_metadata["queue_id"]
    guest = transcript_metadata["guest"].get("jid")
    get_transcript = (
        transcript_metadata["transcript"] or "<div>No transcript found</div>"
    )
    soup = BeautifulSoup(get_transcript, "html.parser")
    divs = soup.find_all("div")
    transcript = list()
    counter = 1
    for div in divs[1::]:
        try:
            transcript.append(
                {
                    "chat_id": chat_id,
                    "message": str(div),
                    "counter": counter,
                    "chat_standalone_url": "https://ca.libraryh3lp.com/dashboard/queues/{0}/calls/REDACTED/{1}".format(
                        queue_id, chat_id
                    ),
                    "guest": guest,
                }
            )
            counter += 1
        except:
            pass
    return transcript


def get_transcript(chat_id):
    """Get the chat info from LibraryH3lp, then retrieve the Transcript out of the Chat.

    Args:
        chat_id (int): A single Chat ID

    Returns:
        list(dict): trancript + metadata
    """
    print("get_transcript")
    client = Client()
    chat_id = int(chat_id)
    transcript_metadata = client.one("chats", chat_id).get()
    transcript = retrieve_transcript(transcript_metadata, chat_id)
    queue_name = transcript_metadata.get("queue").get("name")
    started_date = parse(transcript_metadata.get("started")).strftime("%Y-%m-%d")
    try:
        logging.info("Retrieve transcript for {}".format(str(chat_id)))
        print("Retrieve transcript for {}".format(str(chat_id)))
    except:
        print("error on print")
    return transcript


def get_wait_and_duration(this_chat, started):
    """from the Chat time related metadata...

    Args:
        this_chat (int): Chat metadata

    Returns:
        list: return the Wait Time and Duration Time of the Chat
    """
    print("get_wait_and_duration")
    ended = None
    accepted = None
    wait = None
    duration = None

    try:
        ended = pendulum.parse(this_chat.get("ended"))
    except:
        pass
    try:
        accepted = pendulum.parse(this_chat.get("accepted"))
    except:
        pass
    try:
        wait = accepted - started
    except:
        pass
    try:
        duration = ended - accepted
    except:
        pass

    return [wait, duration]


def get_chat_metadata_for_header(transcript, duration, wait, this_queue):
    """Will generate the section that comes before the transcript on the HTML page.
        It contains metadata information such as Duration and Wait Time of the chat

    Args:
        transcript ([type]): [description]
        duration ([type]): [description]
        wait ([type]): [description]

    Returns:
        string: returning HTML
    """
    print("get_chat_metadata_for_header")

    client = Client()

    if "txt" in this_queue:
        category = "Texting"
    elif "proactive" in this_queue:
        category = "Proactive"
    elif "lavardez" in this_queue:
        category = "Clavardez (fr)"
    else:
        category = "Web"
    
    try:
        duration_in_second = str(datetime.timedelta(0, duration.seconds))
    except:
        duration_in_second = 0
    try:
        wait_in_second = str(datetime.timedelta(0, wait.seconds))
    except:
        wait_in_second = 0

    metadata_html = """
        <div class="container" style="float: right;overflow:hidden;">
            <ul style="overflow:hidden;">
            <li>Duration : <em style="font-weight: 700;">{0}</em> </li>
            <li>Wait : <em style="font-weight: 700;">{1}</em> </li>
            <li><a href="{2}" target="_blank"> {2} </a></li>
            <li><a href="https://forms.office.com/r/DnFj5jjG0h" target="_blank"><em>Form</em></a></li>
            <li>Queue: <em style="font-weight: 700;">{3}</em>  </li>
            <li>Format: <em style="font-weight: 700;">{4}</em>  </li>
            </ul>
        </div>
    """.format(
        duration_in_second,
        wait_in_second,
        transcript[0].get("chat_standalone_url"),
        this_queue,
        category,
    )


    return metadata_html


def line_by_line(transcript, previous_timestamp, operator, html_template, this_queue):
    """Parse each line of the Transcript

    Args:
        transcript (string): [description]
        previous_timestamp (string): [description]
        operator (string): [description]
        html_template (string): [description]
        this_queue (string): [description]

    Returns:
        [type]: HTML Template that will be written to the final file.
    """
    print("line_by_line")
    for message in transcript:
        line_to_add = message.get("message")
        try:
            if previous_timestamp == None:
                previous_timestamp = datetime.time.fromisoformat(
                    (line_to_add[5:10]).strip()
                )
            current_timestamp = datetime.time.fromisoformat((line_to_add[5:10]).strip())
            previous_timestamp = datetime.datetime(
                2011, 11, 11, previous_timestamp.hour, previous_timestamp.minute
            )
            current_timestamp = datetime.datetime(
                2011, 11, 11, current_timestamp.hour, current_timestamp.minute
            )
            timedelta_obj = relativedelta(previous_timestamp, current_timestamp)
        except:
            pass
        if timedelta_obj.minutes >= 5:
            line_to_add = line_to_add.replace(
                current_timestamp,
                """<b style="color:white; background-color:tomato">{0}} [+5] </b>""".format(
                    current_timestamp
                ),
            )
        if operator:
            line_to_add = line_to_add.replace(operator, "operator", 1)
            # import sys; sys.exit()
            # print(line_to_add)
        line_to_add = line_to_add.replace(
            this_queue, "system@chat.ca.libraryh3lp.com", 1
        )
        logging.info("Adding a line")
        html_template.append("<tr><td>" + line_to_add + "</td></tr>")
        #print(html_template)
    return html_template


def generate_html_template_from_transcript(chat_ids, filePath, chat_per_page):
    print("generate_html_template_from_transcript")
    html_template = []
    tracking_guest_id = []
    batch_number = 1

    for index, chat in enumerate(chat_ids):
        try:
            client = Client()
            this_chat = client.one("chats", chat).get()
            
            operator = this_chat.get("operator", {}).get("name", "None")
            if operator:
                pass
            else:
                break
            guest_id = this_chat.get("guest", {}).get("jid")
            chat_id = this_chat.get("guest", {}).get("id")

            tracking_guest_id.append({"guestID": guest_id[0:6], "chat_id": chat_id})

            this_queue = f"{this_chat.get('queue', {}).get('name')}@chat.ca.libraryh3lp.com"
            started = pendulum.parse(this_chat.get("started"))

            wait, duration = get_wait_and_duration(this_chat, started)
            if duration is None:
                continue
            previous_timestamp = None

            transcript = get_transcript(chat)
            metadata_html = get_chat_metadata_for_header(transcript, duration, wait, this_chat.get("queue", {}).get("name"))

            html_template.extend([
                f"<div class='row {'bg-light-blue' if batch_number % 2 == 0 else 'bg-light'}'><h2 class='text-center'> Chat #{index + 1}</h2>",
                metadata_html,
                '<div class="table-responsive"><table class="table mb-0 table-hover">'
            ])
            html_template = line_by_line(transcript, previous_timestamp, operator, html_template, this_queue)
            html_template.append("</table></div></div><div style='background-color: rgb(209, 241, 231); padding:20px 0px; margin:50px 0'></div>")
            
            # Check if it's time to write out this batch
            if (index + 1) % chat_per_page == 0 or (index + 1) == len(chat_ids):
                html = "".join(html_template)
                output = template.render(transcript=html)
                filename = f"{filePath}-{str(uuid4())[:3]}-{batch_number}"

                write_html_to_template(output, f"{filename}.html")
                pd.DataFrame(tracking_guest_id).to_excel(f"{filename}.xlsx")

                # Reset for the next batch
                html_template = []
                tracking_guest_id = []
                batch_number += 1
        except Exception as e:
            print(f"error - generate_html_template_from_transcript: {e}")

def get_chats_for_this_time_range():
    client = Client()
    chats = client.chats()

    chats = chats.list_day(
        year=2022, month=9, day=6, to="2023-08-31"
    )
    return chats


if __name__ == "__main__":
    #Getting chats from a date range
    chats = get_chats_for_this_time_range()
    #Remove anunswered chats
    answered_chats = [chat for chat in chats if chat.get("accepted") is not None]
    #Remove practice chats
    chats = [chat for chat in answered_chats if not "practice" in chat.get("queue")]
    #Transform to DataFrame
    df = pd.DataFrame(chats)
    #Keep only id
    df = df[['id']]

    #If using an excel file from LibraryH3lp Chat History
    # This file contains Unanswered Chats AND Practice Chats also
    #df = pd.read_excel("chatID_for_academic_year_2022-2023.xlsx")
    
    
    sampled_df = df.sample(n=605, random_state=1) # Ensure df is defined and has the 'id' column
    chat_ids = sampled_df['id'].tolist() # Assuming the column name is 'id'
    filePath = "./output/ask"
    chat_per_page = 121
    generate_html_template_from_transcript(chat_ids, filePath, chat_per_page)


