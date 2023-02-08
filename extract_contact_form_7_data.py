from wordpress_xmlrpc import Client, WordPressPost
from wordpress_xmlrpc.methods.posts import GetPosts, NewPost
from wordpress_xmlrpc.methods import media, posts
import pandas as pd
import requests
import os

# Connect to your WordPress site
wp = Client('http://your-site.com/xmlrpc.php', 'your-username', 'your-password')

# Get all Contact Form 7 submissions
posts = wp.call(GetPosts({'post_type': 'wpcf7_contact_form'}))

# Create a list to store the submissions data
submissions = []

# Loop through each submission and get the data
for post in posts:
    submissions.append({
        'subject': post.subject,
        'date': post.date,
        'sender_name': post.sender_name,
        'sender_email': post.sender_email,
        'message': post.message
    })

# Convert the submissions data to a Pandas DataFrame
df = pd.DataFrame(submissions)

# Save the DataFrame to an Excel file
df.to_excel('submissions.xlsx', index=False)


# Get the attachments for each submission
for post in posts:
    # Check if there are any attachments
    if post.attachments:
        # Loop through each attachment
        for attachment in post.attachments:
            # Get the attachment URL
            url = attachment['url']
            
            # Get the attachment file name
            filename = os.path.basename(url)
            
            # Get the attachment type (Lebenslauf, Fotos, or Motivationsschreiben)
            attachment_type = attachment['title']
            
            # Create a directory for the attachment type if it doesn't already exist
            directory = os.path.join('attachments', attachment_type)
            if not os.path.exists(directory):
                os.makedirs(directory)
            
            # Download the attachment
            response = requests.get(url)
            
            # Save the attachment to the appropriate directory
            with open(os.path.join(directory, filename), 'wb') as f:
                f.write(response.content)
