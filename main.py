import pandas as pd
import tableauserverclient as TSC
import smtplib
from email.message import EmailMessage
import logging
import time
import os
import datetime

# Define the log directory path where the files are located
directory_path = "logs/"

# Define the threshold time to delete files that are older than 10 days
threshold_time = datetime.datetime.now() - datetime.timedelta(days=10)

""" SPECIFY SMTP HOST DETAILS BELOW """
smtp_host = 'mail.abc.com'
smtp_port = 25
FROM = 'admin@abc.com'
CC = ''
admin_dl = 'admin@abc.com'

""" SPECIFY TABLEAU SERVER LOGIN DETAILS BELOW """
server_url = 'https://abc.com/'
sites = ''
username = ''
password = ''

LOG_FILE_GEN_TIME = time.strftime("%Y%m%d-%H%M%S")
subscriptions_xl_data = pd.DataFrame()
unlicensed_user_xl_data = pd.DataFrame()

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s, %(levelname)-8s [%(filename)s:%(module)s:%(funcName)s:%(name)s:%(lineno)d] %(message)s',
    datefmt='%Y-%m-%d:%H:%M:%S',
    filename='logs/SubscriptionsRemoval{0}.log'.format(LOG_FILE_GEN_TIME),
    filemode='a'
)

logger = logging.getLogger(__name__)


def delete_subscriptions(all_unlicensed_users):
    global server_url, subscriptions_xl_data
    user_subscriptions_email_data = {}

    try:
        for site_content_url, site_name, user_id, user_name, user_email, last_login, site_id, site_role in all_unlicensed_users:
            logger.info(f'Starting to process user: {user_name} ({site_name})')
            try:
                logger.info(f'Fetching subscriptions for user: {user_name} ({site_name})')
                tableau_auth = TSC.TableauAuth(username, password, site_id=site_content_url)
                server = TSC.Server(server_url)

                try:
                    server.auth.sign_in(tableau_auth)
                    logger.info(f'Authenticated to site: {site_name}')
                except Exception as e:
                    logger.error(f'Failed to authenticate to site: {site_name}. Error: {str(e)}')
                    continue

                try:
                    subscriptions, pagination_item = server.subscriptions.get()
                    logger.info(f'Retrieved {len(subscriptions)} subscriptions for site: {site_name}')
                except Exception as e:
                    logger.error(f'Failed to retrieve subscriptions for site: {site_name}. Error: {str(e)}')
                    continue

                for subscription in subscriptions:
                    if user_id == subscription.user_id:
                        try:
                            view = server.views.get_by_id(subscription.target.id)
                            logger.info(f'Fetched view: {view.name} for subscription: {subscription.id}')
                        except Exception as e:
                            logger.warning(f'Failed to fetch view for subscription: {subscription.id}. Error: {str(e)}')
                            continue

                        if user_email not in user_subscriptions_email_data:
                            user_subscriptions_email_data[user_email] = []

                        if site_name == "Default":
                            link = f"{server_url}#/views/{view.content_url}".replace("sheets/", "")
                        else:
                            link = f"{server_url}#/site/{site_content_url}/views/{view.content_url}".replace("sheets/",
                                                                                                             "")

                        link_html = f'<a href="{link}">Click Here</a>'
                        user_subscriptions_email_data[user_email].append((
                            site_name, user_name, last_login,
                            subscription.subject, view.name, link_html,
                            user_email, site_role
                        ))

                        subscriptions_xl_data = subscriptions_xl_data.append({
                            'SITE_NAME': site_name,
                            'USER_NAME': user_name,
                            'USER_ROLE': site_role,
                            'LAST_LOGIN': last_login,
                            'SUBJECT': subscription.subject,
                            'VIEW_NAME': view.name,
                            'VIEW_URL': link_html
                        }, ignore_index=True)

                        logger.info(
                            f'Processed subscription {subscription.id} for user {user_name} on site {site_name}')
                        # Uncomment the next line to actually delete the subscription
                        server.subscriptions.delete(subscription.id)
                        logger.info(f'Deleted subscription {subscription.id} for view {view.name}')
            except Exception as e:
                logger.error(f'Error processing user {user_name} ({site_name}). Error: {str(e)}')
                continue

    except Exception as e:
        logger.critical(f'Critical error occurred while processing subscriptions. Error: {str(e)}')
    finally:
        logger.info('Finished processing all users')
    subscriptions_xl_data.to_excel('data\\subscriptions_xl_data-failed.xlsx')
    return user_subscriptions_email_data


def get_unlicensed_users():
    global server_url, unlicensed_user_xl_data
    try:
        """ Create connection to Tableau Server """
        logger.info('Connecting to Tableau Server: %s', server_url)
        tableau_auth = TSC.TableauAuth(username, password)
        server = TSC.Server(server_url)
        server.add_http_options({'verify': False})
        """ Loop through sites and fetch unlicensed users """
        all_unlicensed_users = []
        with server.auth.sign_in(tableau_auth):
            logger.info('Connected to Tableau Server: %s', server_url)
            for site in TSC.Pager(server.sites):
                server.auth.switch_site(site)
                logger.info('Authentication switched to site: %s', site.name)
                users, pagination_item = server.users.get()
                for user in users:
                    if user.site_role == 'Unlicensed':
                        logger.info('Fetching unlicensed users for site: %s', site.name)
                        print(f'Fetching unlicensed users for site: {site.name}...')
                        if user.last_login:
                            dt = datetime.datetime.fromisoformat(user.last_login.isoformat())
                            last_login = dt.date()
                        else:
                            last_login = 'Never Logged In'
                        all_unlicensed_users.append((site.content_url, site.name, user.id, user.name, user.email,
                                                     last_login, site.id, user.site_role))
                        """ Generate Dataframe to save data in Excel file """
                        unlicensed_user_xl_data = unlicensed_user_xl_data.append({
                            'SITE_NAME': site.name,
                            'USER_NAME': user.name,
                            'USER_EMAIL': user.email,
                            'USER_ROLE': user.site_role,
                            'LAST_LOGIN': last_login,
                            'USER_FULLNAME': user.fullname,
                            'SITE_CONTENT_URL': site.content_url,
                            'USER_ID': user.id},
                            ignore_index=True)
        logging.info(f"Found unlicensed users: {unlicensed_user_xl_data.to_string()}")
        return all_unlicensed_users
    except Exception as e:
        logging.error(f'Error while deleting subscription...: {str(e)}')


def save_data_excel():
    global unlicensed_user_xl_data, subscriptions_xl_data
    """ Save data to Excel files """
    subscriptions_xl_data.to_excel('data\\subscriptions_xl_data.xlsx')
    unlicensed_user_xl_data.to_excel('data\\unlicensed_user_xl_data.xlsx')


def send_user_email(user_email_id, user_subscriptions):
    global server_url
    df_user_email_data = pd.DataFrame()
    try:
        for subscription in user_subscriptions:
            df_user_email_data = df_user_email_data.append(
                {'SITE_NAME': subscription[0],
                 'USER_NAME': subscription[1],
                 'SITE_ROLE': subscription[7],
                 'LAST_LOGIN': subscription[2],
                 'SUBSCRIPTION_SUBJECT': subscription[3],
                 'SUBSCRIBED_VIEW': subscription[4],
                 'SUBSCRIBED_VIEW_LINK': subscription[5]},
                ignore_index=True)
        msg = EmailMessage()
        msg['Subject'] = 'Your Tableau Subscriptions Have Been Removed'
        msg['From'] = FROM
        msg['To'] = user_email_id
        msg['Bcc'] = 'admin@abc.com'
        row_index = 0
        user_name = df_user_email_data.loc[row_index, 'USER_NAME']
        email_body = df_user_email_data.to_html(na_rep="", index=False, render_links=True,
                                                escape=False).replace('<th>', '<th style="color:white; '
                                                                              'background-color:#180e62">')
        msg.set_content(email_body)
        msg.add_alternative(
            f"""<!DOCTYPE html>
                <html>
                    <body style='font-family: Merriweather; font-size: 11px;'>
                        <p style='color: #44546A;'>Dear <b>{user_name}</b>,</p>
                        <p style='color: #44546A;'>We would like to inform you that your subscriptions from Tableau 
                        Server ({server_url}) have been removed due to your unlicensed status resulting from 
                        inactivity on the server. To regain access to the server, we kindly request that you raise an 
                        access request by following <a 
                        href="https://abc.com</a>. Upon successful grant of access, 
                        you will be able to log in to the server and subscribe to below views.</p> <p style='color: 
                        #44546A;'><b>Details of your removed subscriptions are as follows:</b></p> <br> {email_body} 
                        <br>
                        <p style='color: #44546A;'>Thank you for your understanding and cooperation in this matter.</p>
                        <p style='color: #44546A;'>Regards,<br>Tableau Admin Team</p>
                    </body>
                </html>""",
            subtype='html')
        with smtplib.SMTP(smtp_host, smtp_port) as smtp:
            smtp.send_message(msg)
            logger.info('Email sent to user: %s', user_email_id)
    except Exception as e:
        logging.error(f"Error sending email: {str(e)}")


def send_success_email_admin():
    global server_url, subscriptions_xl_data
    try:
        msg = EmailMessage()
        msg['Subject'] = "Success - Tableau Subscriptions Cleanup Operation Completed"
        msg['From'] = FROM
        msg['To'] = admin_dl
        email_body = subscriptions_xl_data.to_html(na_rep="", index=False, render_links=True,
                                                   escape=False).replace('<th>', '<th style="color:white; '
                                                                                 'background-color:#180e62">')
        msg.set_content(email_body)
        msg.add_alternative(
            f"""<!DOCTYPE html>
                <html>
                    <body style='font-family: Merriweather; font-size: 11px;'>
                        <p style='color: #44546A;'>Hi Team,</p>
                        <p style='color: #44546A;'>Subscriptions cleanup activity completed on Tableau Server ({server_url}).</p>
                        <p style='color: #44546A;'><b>Details of all removed subscriptions are as follows:</b></p>
                        <br>
                        {email_body}
                        <br>
                        <p style='color: #44546A;'>Regards,<br>Subscription Automation</p>
                    </body>
                </html>""",
            subtype='html'
        )
        with smtplib.SMTP(smtp_host, smtp_port) as smtp:
            smtp.send_message(msg)
            logger.info('Email sent to Tableau Admin Team For Completion.')
    except Exception as e:
        logging.error(f"Error sending email to Tableau Admin Team: {str(e)}")


def send_failed_email_admin():
    global server_url
    try:
        msg = EmailMessage()
        msg['Subject'] = "Failed - Tableau Subscriptions Cleanup Operation"
        msg['From'] = FROM
        msg['To'] = admin_dl
        msg.set_content('')
        msg.add_alternative(
            f"""<!DOCTYPE html>
                <html>
                    <body style='font-family: Merriweather; font-size: 11px;'>
                        <p style='color: #44546A;'>Hi Team,</p>
                        <p style='color: #44546A;'>Subscriptions cleanup activity failed on Tableau Server ({server_url}), 
                        please check the server logs to investigate the issue.</p>
                        <p style='color: #44546A;'>Regards,<br>Subscription Automation</p>
                    </body>
                </html>""",
            subtype='html'
        )
        with smtplib.SMTP(smtp_host, smtp_port) as smtp:
            smtp.send_message(msg)
            logger.info('Email sent to Tableau Admin Team For Failure.')
    except Exception as e:
        logging.error(f"Error sending email to Tableau Admin Team For Failed Operation: {str(e)}")


def delete_logs():
    global directory_path
    # Loop through all files in the logs directory
    try:
        for file_name in os.listdir(directory_path):
            # Get the creation time of the file
            file_path = os.path.join(directory_path, file_name)
            creation_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path))

            # Check if the file is older than the threshold time
            if creation_time < threshold_time:
                # Delete the file if it's older than the threshold time
                os.remove(file_path)
                print(f"Log Deleted file: {file_name}")
                logger.info(f"Log file older than 10 days deleted file: {file_name}")
    except Exception as e:
        logging.error(f"Error while deleting log files: {str(e)}")


if __name__ == '__main__':
    logging.info(('#' * 15) + ' Subscription CleanUp Operation Started ' + ('#' * 15))
    try:
        delete_logs()
        returned_all_unlicensed_users = get_unlicensed_users()
        user_email_data = delete_subscriptions(returned_all_unlicensed_users)
        """ Send list of subscriptions removed to each user """
        if user_email_data:
            for user_email, subscriptions in user_email_data.items():
                send_user_email(user_email, subscriptions)
                logger.info(f'Successfully sent email to : {user_email}...')
                print(f'Successfully sent email to : {user_email}')
                logger.info(f'Sent Subscription Data : {str(subscriptions)}')
            send_success_email_admin()
            save_data_excel()
            logging.info(('#' * 15) + ' Subscription CleanUp Operation Completed ' + ('#' * 15))
            print("Done!")
    except Exception as e:
        send_failed_email_admin()
        logging.error(f"Error in __main__ function : {str(e)}")
        logging.debug(('#' * 15) + ' Subscription CleanUp Operation Failed ' + ('#' * 15))
