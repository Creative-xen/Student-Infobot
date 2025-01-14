import pandas as pd
from telegram import Update, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# Read the Excel sheet - telling to pandas to read it 
excel_file = 'your_data.xlsx'  
df = pd.read_excel(excel_file)

# roll are string 
df['roll'] = df['roll'].astype(str)


user_data_file = 'user_data.xlsx'


def read_or_initialize_user_data():
    try:
        return pd.read_excel(user_data_file)
    except FileNotFoundError:
        return pd.DataFrame(columns=['user_id', 'username'])


def save_user_data(user_data_df):
    user_data_df.to_excel(user_data_file, index=False)


user_data = read_or_initialize_user_data()


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    username = update.effective_user.username
    
    if username:
        username = f"@{username}"
    else:
        username = "N/A"
    
    global user_data
    if user_data[(user_data['user_id'] == user_id) & (user_data['username'] == username)].empty:
        new_user = pd.DataFrame([[user_id, username]], columns=['user_id', 'username'])
        user_data = pd.concat([user_data, new_user], ignore_index=True)
        save_user_data(user_data)
    
    await help_command(update, context)

                                                                                                                                      #formalaties :(
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "Hardik Swagat to the Student Info Bot!\n\n"
        "Here are the commands you can use:\n\n"
        "/start - Get a welcome message.\n\n"
        
        "To get details about a student, send a roll number.\nFor example: '220517xx'.\n\n"
        "To get a list of students in a section, send a section number in the format '01' for CSE-01.\n"
        "I will return the name, roll number, and hostel for each student in that section, sorted by roll number.\n\n\n"
        "Developed by :  Ankush"

    )

    await update.message.reply_text(help_text)


async def handle_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    
                                                                                                                                # Check if the command is a roll number (5-9 digits)
    if query.isdigit() and 5 <= len(query) <= 9:
        await get_data(update, context, query)
                                                                                                                                 # Check if the command is a section number (1-2 digits)
    elif query.isdigit() and 1 <= len(query) <= 2:
        await get_section(update, context, query)
    else:
        await update.message.reply_text("Please enter a valid roll number (5-9 digits) or section number (1-2 digits).")


async def get_data(update: Update, context: ContextTypes.DEFAULT_TYPE, roll_number: str):
    student_data = df[df['roll'] == roll_number]
    
    if not student_data.empty:
        row = student_data.iloc[0]
        message = (f"Name - {row['name']}\n"
                   f"Roll No - {row['roll']}\n"
                   f"Section - {row['section']}\n"
                   f"Hostel - {row['hostel']}")
    else:
        message = "Roll number not found."

    await update.message.reply_text(message)


async def get_section(update: Update, context: ContextTypes.DEFAULT_TYPE, section: str):
    section_data = df[df['section'].str.contains(f"CSE-{section.zfill(2)}", na=False)]
    
    if not section_data.empty:
        section_data_sorted = section_data.sort_values(by='roll')
        
                                                                                                                   # break if mess is long to post
        message_parts = [
            f"Name - {row['name']}\nRoll No - {row['roll']}\nHostel - {row['hostel']}\n"
            for _, row in section_data_sorted.iterrows()
        ]
        
        
        chunk_size = 30                                                                                              #  chunk size as needed choose any (helps to acheive max word limit from tele)
        for i in range(0, len(message_parts), chunk_size):
            chunk = "\n".join(message_parts[i:i+chunk_size])
            await update.message.reply_text(chunk)
    else:
        await update.message.reply_text("Section not found.")

                                                                                                                     # Access to handle members command ( only me ) 
async def users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = 1999878201
    if update.effective_user.id == admin_id:
        file_path = 'user_data.xlsx'
        user_data.to_excel(file_path, index=False)
        
        with open(file_path, 'rb') as file:
            await update.message.reply_document(document=InputFile(file, filename='user_data.xlsx'))
    else:
        await update.message.reply_text("You are not authorized to use this command.")


def main():
                                                                                                                     # Replace 'TOKEN' put yours here
    application = Application.builder().token("<your bot token>").build()

  
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("users", users))
    

    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_query))  

    application.run_polling() 

if __name__ == "__main__":
    main()
