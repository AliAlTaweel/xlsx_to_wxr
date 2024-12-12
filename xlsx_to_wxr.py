import openpyxl
from datetime import datetime

def convert_xlsx_to_wxr(xlsx_path, wxr_output_path):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(xlsx_path)
    sheet = workbook.active

    # Start WXR file content
    wxr_content = """<?xml version="1.0" encoding="UTF-8" ?>
    <rss version="2.0" xmlns:excerpt="http://wordpress.org/export/1.2/excerpt/"
        xmlns:content="http://purl.org/rss/1.0/modules/content/"
        xmlns:wfw="http://wellformedweb.org/CommentAPI/"
        xmlns:dc="http://purl.org/dc/elements/1.1/"
        xmlns:wp="http://wordpress.org/export/1.2/">
    <channel>
    <title>My WordPress Events</title>
    <link>https://example.com</link>
    <description>Events imported from Excel</description>
    <language>en</language>
    <wp:wxr_version>1.2</wp:wxr_version>
    """

    # Loop through Excel rows (assuming first row is the header)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        event_start_date = row[0]  # Aika (Event Date)
        event_name = row[1]  # Tapahtuma (Event Title)
        event_organizer = row[2]  # Vastuulliset (Responsible)
        mentor_name = row[3]  # Mentor
        post_content = row[4]  # Kuvaus (Description)
        event_place = row[5]  # Paikka (Place/Category)
        event_author = row[6]  # Author

        # Validate the event start date
        if isinstance(event_start_date, datetime):
            event_start_date = event_start_date.strftime('%Y-%m-%d')  # Format as YYYY-MM-DD
        else:
            event_start_date = "-"  # Default date if invalid or missing

        # Ensure the event title is a string and not empty
        if not event_name:
            event_name = f"Untitled Event {row[0]}"  # Default title if missing
        event_name = str(event_name)

        # Convert the event title to a post name (slug)
        post_name = event_name.lower().replace(' ', '-')

        # Default the author name if missing
        if not mentor_name:
            mentor_name = "-"

        # Ensure event_place is not empty
        event_place = event_place if event_place else "Ei Paikka"

        # Add the event to the WXR content
        wxr_content += f"""
        <item>
            <title>{event_name}</title>
            <wp:post_date>{event_start_date} 00:00:00</wp:post_date>
            <wp:post_name>{post_name}</wp:post_name>
            <wp:post_author>{ event_author}</wp:post_author>
            <wp:post_type>event</wp:post_type>
            <wp:status>publish</wp:status>
            <content:encoded><![CDATA[{post_content}]]></content:encoded>
            <wp:postmeta>
                <wp:meta_key>_event_start_date</wp:meta_key>
                <wp:meta_value>{event_start_date}</wp:meta_value>
            </wp:postmeta>
            <wp:postmeta>
                <wp:meta_key>_event_end_date</wp:meta_key>
                <wp:meta_value>{event_start_date}</wp:meta_value>  <!-- Assuming end date is the same -->
            </wp:postmeta>
            <wp:postmeta>
                <wp:meta_key>_responsible_name</wp:meta_key>
                <wp:meta_value>{event_organizer}</wp:meta_value>
            </wp:postmeta>
            <wp:postmeta>
                <wp:meta_key>_mentor_names</wp:meta_key>
                <wp:meta_value>{mentor_name}</wp:meta_value>
            </wp:postmeta>
            <category domain="Place" nicename="{event_place.lower().replace(' ', '-')}">
                <![CDATA[{event_place}]]>
            </category>
        </item>
        """

    # End WXR file content
    wxr_content += """
    </channel>
    </rss>
    """

    # Write the WXR file
    with open(wxr_output_path, "w", encoding="utf-8") as file:
        file.write(wxr_content)

    print(f"WXR file created: {wxr_output_path}")

# Define the file paths
xlsx_path = "./events 2024-2018.xlsx"  # Update this to your Excel file path
wxr_output_path = "output_file.xml"  # Update this to your desired output file path

# Convert the Excel file to WXR
convert_xlsx_to_wxr(xlsx_path, wxr_output_path)
