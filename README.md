# sendEmail

The program searches for PDF files in a special naming format (group_name+attach1.....pdf) and sends them to different email addresses.

The addresses are loaded from a CSV file containing the group name and the email addresses.

## File name format of the PDF files

The two spaces (also after the Firstname) are important, because they are referenced for the search. All stuff behind is not important. Only the .pdf extension.

Groupname           additional Stuff .pdf

Examples:

```text
Lastname Firstname xxxxxxxxxxxx xx x x.pdf
Smith Kaith bill.pdf
Gilles Tom Bill.pdf
Smith Kaith incoice.pdf
Smith Keith chocolate.pdf
Gilles Tom cake.pdf
Mueller Craig chocolate cake.pdf
```

## To do:

* The difficult part is to create the Google MAIL API Json file according to the video [Getting Started With Google APIs For Python Development from Jie Jenn](https://www.youtube.com/watch?v=PKLG5pfs4nY&t=0s). IMPORTANT: Activate Google Mail API

### In the main folder

* Store the `client_secret.json` file (from the Video above) in the main folder
* Rename and edit the `.env.example` file to `.env` with the appropriate values 
* Create and activate a venv and load the modules according to the `requirements.txt` file.

### In the config folder
* Rename the `groups.csv.example` file to `groups.csv` and fill in the working groups and emails.
* Rename and edit the `mail.txt.example` file to `mail.txt` and insert the html format body text.   

## Start of the program

Now you can start the program in the venv:

```bash
python3 sendEmail.py
```

The program goes through the PDF directory, searches acording to the groups.csv and send the emails. After sending of each email the related PDF files are moved to the old directory. 
