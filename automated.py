import pandas as pd
import numpy as np
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass

# log into email via SMTP
print("\nHi! Please log in.\n")
me = input("Email address: ")
pw = getpass("Password (keys hidden): ")

mail = smtplib.SMTP('smtp.gmail.com', 587)

mail.ehlo()

mail.starttls()

mail.login(me, pw)

my_name = input("First name (this is for your greeting & signature): ")
my_sheet = input("Sheet name (exactly as it appears & with .xlsx at end ): ")

# note - donor email cells must be sorted into abc order before reading
df = pd.read_excel(my_sheet)

# turn names column into numpy array
donors = df["donor email"].to_numpy()

# list to collect matches missing a donor or recipient email
not_sent = []

# list to collect matches that were sent
sent_list = []


def email():
    '''
    goes through match sheet and sends the emails

            parameters:
                none

            returns:
                none
    '''
    clean_df()
    while donors.size != 0:
        if donors.size == 1:
            single(df)
        else:
            chunk = df[df["donor email"] == donors[0]].copy(deep=True)

            if donors[0] == donors[1]:
                multi(chunk)
            else:
                single(chunk)
        cut()
    send_results()
    print("Finished!")


def clean_df():
    '''
    drops unncessary columns from df

            parameters:
                none

            returns:
                none
    '''
    df.drop(['donor name', 'donor type', 'donor capacity', 'donor open_box',
             'donor zip_region', 'donor phone', 'donor contact_method',
             'donor address', 'donor state', 'donor city',
             'donor address1', 'donor address2', 'donor transfer_types',
             'donor anything_else', 'donor other_ppe', 'donor zip',
             'recipient name', 'recipient capacity', 'recipient open_box',
             'recipient facility_type', 'recipient current_class',
             'recipient priority', 'recipient contact_title',
             'recipient address', 'recipient state', 'recipient city',
             'recipient address2', 'recipient loading_dock',
             'recipient drop_instructions', 'recipient anything_else',
             'recipient zip', 'distance'], inplace=True, axis=1)


def multi(d):
    '''
    sends the multi-recipient emails

            parameters:
                    d (dataframe): subsection of match sheet dataframe, df,
                        that matches donor email

            returns:
                    none
    '''
    donor = donors[0]
    idx = np.flatnonzero(df["donor email"] == donor)[0]
    nm = d.loc[idx, "donor contact_name"]

    if pd.isna(donor):
        not_sent.append(nm)
    else:
        # message container
        msg = MIMEMultipart('alternative')
        msg['Subject'] = "We’ve matched your PPE donation to several nearby \
            organizations in need"
        msg['From'] = me
        msg['To'] = donor

        table = stylize_frame(d)

        # message body
        html = (
            "<html>"
            "<head></head>"
            "<body>"
            f"<p>Hi {nm.split()[0]}!</p>"
            f"<p>This is {my_name} writing from"
            " <a href = GetUsPPE.org > GetUsPPE.org </a> -"
            " we’re incredibly grateful that you’ve offered to donate "
            "supplies to healthcare organizations on the front lines, and "
            "I’m happy to say that we’ve found several high-need "
            "organizations in your area that could benefit from your"
            "donation.</p>"
            f"<p>We would like to split your donations so that they’re "
            f"distributed to the organizations below:<br> {table} <br>"
            "Would you be able to deliver/ship to all of these "
            "organizations in need? <b> If not, please respond to this "
            "email indicating that you would like volunteer assistance with "
            "delivering PPE to these various organizations. </b></p>"
            "<p>Thanks so much for your time and consideration,<br>"
            f"{my_name} <br>"
            "National GetUsPPE Volunteer<br>"
            "<a href = match@getusppe.org > match@getusppe.org </a></p>"
            "<p>--</p>"
            "<p>GetUsPPE <b>Match Team</b> - match@GetUsPPE.org</p>"
            "</body>"
            "</html>"
        )

        # record MIME type
        body = MIMEText(html, 'html')

        # attach text into message container
        msg.attach(body)

        try:
            mail.sendmail(me, donor, msg.as_string())

            sent_list.append(f"Multi: {nm}: {donor}")

            print(f"Sending email to {nm}: {donor}..... (multi)")

        except smtplib.SMTPRecipientsRefused:

            not_sent.append(nm)


def stylize_frame(d):
    '''
    formats the dataframe subsection and converts it to html so that it can be
    pasted into the email as a table

            parameters:
                    d (dataframe): subsection of match sheet dataframe, df,
                    that matches donor email

            returns:
                    d converted to an html table
    '''
    d.drop(["donor contact_name", "donor email", "recipient contact_name",
            "recipient email", "recipient phone"],
           axis=1, inplace=True)

    d.columns = ["Donation Quantity", "PPE Type", "Recipient Facility Name",
                 "Recipient Address"]

    d = d[["Recipient Facility Name", "Recipient Address", "PPE Type",
           "Donation Quantity"]]

    d["PPE Type"] = d["PPE Type"].str.replace("_", " ")

    d.fillna("N/A", inplace=True)

    d["Donation Quantity"] = d["Donation Quantity"].astype(int)

    return d.to_html(index=False, justify="center")


def single(d):
    '''
    sends the single-recipient emails

            parameters:
                    d (dataframe): subsection of match sheet dataframe, df,
                        that matches donor email

            returns:
                    none
    '''
    donor = donors[0]
    idx = np.flatnonzero(df["donor email"] == donor)[0]
    d_nm = d.loc[idx, "donor contact_name"]
    recipient = d.loc[idx, "recipient email"]

    if pd.isna(donor) or pd.isna(recipient):
        not_sent.append(d_nm)
    else:
        d.fillna("N/A", inplace=True)
        r_nm = d.loc[idx, "recipient contact_name"]
        r_type = d.loc[idx, "recipient type"]
        r_org = d.loc[idx, "recipient facility_name"]
        r_address = d.loc[idx, "recipient address1"]
        r_phone = d.loc[idx, "recipient phone"]

        if "_" in r_type:
            r_type = " ".join(r_type.split("_"))

        # message container
        msg = MIMEMultipart('alternative')
        msg['Subject'] = "We’ve matched your PPE donation to a nearby \
            organization in need"
        msg['From'] = "GetUsPPE Match - match@getusppe.org" + ' <' + me + '>'
        msg['Reply-to'] = "math.reading.science@gmail.com"
        msg['To'] = donor
        msg['Cc'] = recipient

        # message body
        html = (
            "<html>"
            "<head></head>"
            "<body>"
            f"<p>Hi {d_nm.split()[0]}!</p>"
            "<p>I'm a volunteer writing from "
            "<a href = GetUsPPE.org > GetUsPPE.org </a> - we’re incredibly "
            "grateful that you've offered to donate supplies to healthcare "
            "organizations on the front lines, and I'm happy to say that "
            "we've found an organization in your area in high need! </p>"
            "<p>The most urgent need in the GetUsPPE database for "
            f"{r_type} in your area is <b>{r_org}</b>. A contact person "
            " their organization has been copied on this email. </p>"
            f"<p>Their address/drop-off location is: {r_address} </p>"
            "<p>And, if needed their contact info is: </p>"
            f"<p>Name: {r_nm} <br>"
            f"Phone: {r_phone} <br>"
            f"Email: check CC </p>"
            "<p><u>Recipients</u>: Please reply all indicating if this "
            "PPE is still needed at your location and whether the included "
            "drop-off/delivery address is correct. We also want to note "
            "that many donors have been seeking multiple avenues to donate "
            "PPE as quickly as possible, so some of these matched "
            "donations may have already been given elsewhere. "
            "Importantly, <b> we ask that you respond to this email "
            "letting us know once you have received this donation</b>; "
            "you can do so by replying to this email. We understand that "
            "this is one extra step on your end and greatly appreciate it "
            "- this is truly a situation where data, like PPE, saves "
            "lives!</p>"
            "<p><u>Donors</u>: Please reply, including both GetUsPPE and "
            "the copied recipient, if you have already donated this PPE "
            "elsewhere and it is no longer available. In addition, please "
            "ensure that the recipient organization is still able to "
            "accept donations, as needs and circumstances occasionally "
            "change quickly for our requesting organizations. Remember, "
            "you can always ship donations! Finally, <b> if you do plan "
            "to donate but it turns out that you are unable to find a way "
            "to deliver/ship your donation, we may be able to find a "
            "local volunteer in your area who can help transport it, "
            "please reply to this email informing us of that need.</b></p>"
            "<p>We are incredibly grateful to all of you who are either "
            "caring for patients or supporting frontline healthcare "
            "workers during this difficult time through these donations. "
            "Please let us know if you have any questions, and thank you "
            "for all that you do. </p>"
            "</p>Sincerely,<br>"
            f"{my_name} <br>"
            "National GetUsPPE Volunteer<br>"
            "<a href = match@getusppe.org > match@getusppe.org </a></p>"
            "<p> -- </p>"
            "<p>GetUsPPE <b>Match Team</b> - match@GetUsPPE.org</p>"
            "</body>"
            "</html>"
        )

        # record MIME type
        body = MIMEText(html, 'html')

        # attach text into message container
        msg.attach(body)

        try:
            mail.sendmail(me, [donor, recipient], msg.as_string())

            sent_list.append(f"Single: {d_nm}: {donor} to {r_nm}: {recipient}")

            print(
                f"Sending email to {d_nm}: {donor} and {r_nm}: {recipient}....."
                "(single)")

        except smtplib.SMTPRecipientsRefused:

            not_sent.append(d_nm)


def cut():
    '''
    deletes the first n consecutive, identical donor emails arr

            parameters:
                none

            returns:
                none
    '''
    global donors
    hd = donors[0]
    for i, nm in enumerate(donors):
        if nm != hd:
            break
    if donors[0] == donors[-1]:
        donors = np.empty(0)
    else:
        donors = donors[i:]


def send_results():
    # message container
    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"Match Sheet Results for {my_name}"
    msg['From'] = me
    msg['To'] = me

    html = (
        "<html>"
        "<head></head>"
        "<body>"
        f"<p>Hi {my_name}!</p>"
        "<p>These are your results:</p>"
        f"<p>{results()}</p>"
        "</body>"
        "</html>"
    )

    body = MIMEText(html, "html")
    msg.attach(body)
    mail.sendmail(me, me, msg.as_string())


def results():
    '''
    prints unsent emails

            parameters:
                none

            returns:
                none
    '''
    sent = "<br>".join(sent_list)
    failed = "<br>".join(not_sent)

    if not_sent:
        r = (f"<p><em>Sent:</em></p> <p> {sent}</p> <p><em>Missing donor or "
             f"recipient email for:</em></p> <p> {failed}</p>")
    else:
        r = f"<p>Sent:</p> <p> {sent}</p>"

    return r


def main():
    email()


if __name__ == "__main__":
    main()
