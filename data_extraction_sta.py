#data extraction
import requests
from bs4 import BeautifulSoup
import time
import os

# Function to extract article text
def extract_article_text(url):
    try:
        # Make a request to the website
        r = requests.get(url)
        r.raise_for_status()

        # Parse the HTML content
        soup = BeautifulSoup(r.content, 'html.parser')

        # Find the article text
        article_title = soup.find('h1', {'class': 'entry-title'}) or soup.find('h1', {'class': 'tdb-title-text'})
        article_text = soup.find('div', {'class': 'td_block_wrap tdb_single_content tdi_130 td-pb-border-top td_block_template_1 td-post-content tagdiv-type'}) or soup.find('div', {'class': 'td-post-content tagdiv-type'})

        # Exempt content from pre class = "wp-block-preformatted"
        preformatted_text = article_text.find('pre', {'class': 'wp-block-preformatted'})
        if preformatted_text:
            preformatted_text.decompose()

        # Return the article title and text
        return article_title.text, article_text.text
    except requests.exceptions.HTTPError as errh:
        print("Ooops... Error 404\nSorry, but the page you are looking for doesn't exist.")
        return None, None
    except requests.exceptions.ConnectionError as errc:
        print("Ooops... Connection Error\nPlease check your internet connection and try again.")
        return None, None
    except requests.exceptions.Timeout as errt:
        print("Ooops... Timeout Error\nThe request took too long to complete.")
        return None, None
    except requests.exceptions.RequestException as err:
        print("Ooops... Something went wrong \n", err)
        return None, None

# list of 100 articles URLs
urls = ['https://insights.blackcoffer.com/rising-it-cities-and-its-impact-on-the-economy-environment-infrastructure-and-city-life-by-the-year-2040-2/',
'https://insights.blackcoffer.com/rising-it-cities-and-their-impact-on-the-economy-environment-infrastructure-and-city-life-in-future/',
'https://insights.blackcoffer.com/internet-demands-evolution-communication-impact-and-2035s-alternative-pathways/',
'https://insights.blackcoffer.com/rise-of-cybercrime-and-its-effect-in-upcoming-future/',
'https://insights.blackcoffer.com/ott-platform-and-its-impact-on-the-entertainment-industry-in-future/',
'https://insights.blackcoffer.com/the-rise-of-the-ott-platform-and-its-impact-on-the-entertainment-industry-by-2040/',
'https://insights.blackcoffer.com/rise-of-cyber-crime-and-its-effects/',
'https://insights.blackcoffer.com/rise-of-internet-demand-and-its-impact-on-communications-and-alternatives-by-the-year-2035-2/',
'https://insights.blackcoffer.com/rise-of-cybercrime-and-its-effect-by-the-year-2040-2/',
'https://insights.blackcoffer.com/rise-of-cybercrime-and-its-effect-by-the-year-2040/',
'https://insights.blackcoffer.com/rise-of-internet-demand-and-its-impact-on-communications-and-alternatives-by-the-year-2035/',
'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-3-2/',
'https://insights.blackcoffer.com/rise-of-e-health-and-its-impact-on-humans-by-the-year-2030/',
'https://insights.blackcoffer.com/rise-of-e-health-and-its-imapct-on-humans-by-the-year-2030-2/',
'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-2/',
'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-2-2/',
'https://insights.blackcoffer.com/rise-of-chatbots-and-its-impact-on-customer-support-by-the-year-2040/',
'https://insights.blackcoffer.com/rise-of-e-health-and-its-imapct-on-humans-by-the-year-2030/',
'https://insights.blackcoffer.com/how-does-marketing-influence-businesses-and-consumers/',
'https://insights.blackcoffer.com/how-advertisement-increase-your-market-value/',
'https://insights.blackcoffer.com/negative-effects-of-marketing-on-society/',
'https://insights.blackcoffer.com/how-advertisement-marketing-affects-business/',
'https://insights.blackcoffer.com/rising-it-cities-will-impact-the-economy-environment-infrastructure-and-city-life-by-the-year-2035/',
'https://insights.blackcoffer.com/rise-of-ott-platform-and-its-impact-on-entertainment-industry-by-the-year-2030/',
'https://insights.blackcoffer.com/rise-of-electric-vehicles-and-its-impact-on-livelihood-by-2040/',
'https://insights.blackcoffer.com/rise-of-electric-vehicle-and-its-impact-on-livelihood-by-the-year-2040/',
'https://insights.blackcoffer.com/oil-prices-by-the-year-2040-and-how-it-will-impact-the-world-economy/',
'https://insights.blackcoffer.com/an-outlook-of-healthcare-by-the-year-2040-and-how-it-will-impact-human-lives/',
'https://insights.blackcoffer.com/ai-in-healthcare-to-improve-patient-outcomes/',
'https://insights.blackcoffer.com/what-if-the-creation-is-taking-over-the-creator/',
'https://insights.blackcoffer.com/what-jobs-will-robots-take-from-humans-in-the-future/',
'https://insights.blackcoffer.com/will-machine-replace-the-human-in-the-future-of-work/',
'https://insights.blackcoffer.com/will-ai-replace-us-or-work-with-us/',
'https://insights.blackcoffer.com/man-and-machines-together-machines-are-more-diligent-than-humans-blackcoffe/',
'https://insights.blackcoffer.com/in-future-or-in-upcoming-years-humans-and-machines-are-going-to-work-together-in-every-field-of-work/',
'https://insights.blackcoffer.com/how-neural-networks-can-be-applied-in-various-areas-in-the-future/',
'https://insights.blackcoffer.com/how-machine-learning-will-affect-your-business/',
'https://insights.blackcoffer.com/deep-learning-impact-on-areas-of-e-learning/',
'https://insights.blackcoffer.com/how-to-protect-future-data-and-its-privacy-blackcoffer/',
'https://insights.blackcoffer.com/how-machines-ai-automations-and-robo-human-are-effective-in-finance-and-banking/',
'https://insights.blackcoffer.com/ai-human-robotics-machine-future-planet-blackcoffer-thinking-jobs-workplace/',
'https://insights.blackcoffer.com/how-ai-will-change-the-world-blackcoffer/',
'https://insights.blackcoffer.com/future-of-work-how-ai-has-entered-the-workplace/',
'https://insights.blackcoffer.com/ai-tool-alexa-google-assistant-finance-banking-tool-future/',
'https://insights.blackcoffer.com/ai-healthcare-revolution-ml-technology-algorithm-google-analytics-industrialrevolution/',
'https://insights.blackcoffer.com/all-you-need-to-know-about-online-marketing/',
'https://insights.blackcoffer.com/evolution-of-advertising-industry/',
'https://insights.blackcoffer.com/how-data-analytics-can-help-your-business-respond-to-the-impact-of-covid-19/',
'https://insights.blackcoffer.com/covid-19-environmental-impact-for-the-future/',
'https://insights.blackcoffer.com/environmental-impact-of-the-covid-19-pandemic-lesson-for-the-future/',
'https://insights.blackcoffer.com/how-data-analytics-and-ai-are-used-to-halt-the-covid-19-pandemic/',
'https://insights.blackcoffer.com/difference-between-artificial-intelligence-machine-learning-statistics-and-data-mining/',
'https://insights.blackcoffer.com/how-python-became-the-first-choice-for-data-science/',
'https://insights.blackcoffer.com/how-google-fit-measure-heart-and-respiratory-rates-using-a-phone/',
'https://insights.blackcoffer.com/what-is-the-future-of-mobile-apps/',
'https://insights.blackcoffer.com/impact-of-ai-in-health-and-medicine/',
'https://insights.blackcoffer.com/telemedicine-what-patients-like-and-dislike-about-it/',
'https://insights.blackcoffer.com/how-we-forecast-future-technologies/',
'https://insights.blackcoffer.com/can-robots-tackle-late-life-loneliness/',
'https://insights.blackcoffer.com/embedding-care-robots-into-society-socio-technical-considerations/',
'https://insights.blackcoffer.com/management-challenges-for-future-digitalization-of-healthcare-services/',
'https://insights.blackcoffer.com/are-we-any-closer-to-preventing-a-nuclear-holocaust/',
'https://insights.blackcoffer.com/will-technology-eliminate-the-need-for-animal-testing-in-drug-development/',
'https://insights.blackcoffer.com/will-we-ever-understand-the-nature-of-consciousness/',
'https://insights.blackcoffer.com/will-we-ever-colonize-outer-space/',
'https://insights.blackcoffer.com/what-is-the-chance-homo-sapiens-will-survive-for-the-next-500-years/',
'https://insights.blackcoffer.com/why-does-your-business-need-a-chatbot/',
'https://insights.blackcoffer.com/how-you-lead-a-project-or-a-team-without-any-technical-expertise/',
'https://insights.blackcoffer.com/can-you-be-great-leader-without-technical-expertise/',
'https://insights.blackcoffer.com/how-does-artificial-intelligence-affect-the-environment/',
'https://insights.blackcoffer.com/how-to-overcome-your-fear-of-making-mistakes-2/',
'https://insights.blackcoffer.com/is-perfection-the-greatest-enemy-of-productivity/',
'https://insights.blackcoffer.com/global-financial-crisis-2008-causes-effects-and-its-solution/',
'https://insights.blackcoffer.com/gender-diversity-and-equality-in-the-tech-industry/',
'https://insights.blackcoffer.com/how-to-overcome-your-fear-of-making-mistakes/',
'https://insights.blackcoffer.com/how-small-business-can-survive-the-coronavirus-crisis/',
'https://insights.blackcoffer.com/impacts-of-covid-19-on-vegetable-vendors-and-food-stalls/',
'https://insights.blackcoffer.com/impacts-of-covid-19-on-vegetable-vendors/',
'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-tourism-aviation-industries/',
'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-sports-events-around-the-world/',
'https://insights.blackcoffer.com/changing-landscape-and-emerging-trends-in-the-indian-it-ites-industry/',
'https://insights.blackcoffer.com/online-gaming-adolescent-online-gaming-effects-demotivated-depression-musculoskeletal-and-psychosomatic-symptoms/',
'https://insights.blackcoffer.com/human-rights-outlook/',
'https://insights.blackcoffer.com/how-voice-search-makes-your-business-a-successful-business/',
'https://insights.blackcoffer.com/how-the-covid-19-crisis-is-redefining-jobs-and-services/',
'https://insights.blackcoffer.com/how-to-increase-social-media-engagement-for-marketers/',
'https://insights.blackcoffer.com/impacts-of-covid-19-on-streets-sides-food-stalls/',
'https://insights.blackcoffer.com/coronavirus-impact-on-energy-markets-2/',
'https://insights.blackcoffer.com/coronavirus-impact-on-the-hospitality-industry-5/',
'https://insights.blackcoffer.com/lessons-from-the-past-some-key-learnings-relevant-to-the-coronavirus-crisis-4/',
'https://insights.blackcoffer.com/estimating-the-impact-of-covid-19-on-the-world-of-work-2/',
'https://insights.blackcoffer.com/estimating-the-impact-of-covid-19-on-the-world-of-work-3/',
'https://insights.blackcoffer.com/travel-and-tourism-outlook/',
'https://insights.blackcoffer.com/gaming-disorder-and-effects-of-gaming-on-health/',
'https://insights.blackcoffer.com/what-is-the-repercussion-of-the-environment-due-to-the-covid-19-pandemic-situation/',
'https://insights.blackcoffer.com/what-is-the-repercussion-of-the-environment-due-to-the-covid-19-pandemic-situation-2/',
'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-office-space-and-co-working-industries/',
'https://insights.blackcoffer.com/contribution-of-handicrafts-visual-arts-literature-in-the-indian-economy/',
'https://insights.blackcoffer.com/how-covid-19-is-impacting-payment-preferences/',
'https://insights.blackcoffer.com/how-will-covid-19-affect-the-world-of-work-2/']

# Create a directory on the desktop
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
directory_path = os.path.join(desktop_path, "blackassign_articles")

# Check if the directory already exists, if not, create it
if not os.path.exists(directory_path):
    os.makedirs(directory_path)

# Loop through the URLs and extract the article text
for i, url in enumerate(urls):
    article_title, article_text = extract_article_text(url)

    if article_title and article_text:
        # Define the file name
        file_name = f'blackassign{str(i + 1).zfill(4)}.txt'
        
        # Create the full path to the file
        file_path = os.path.join(directory_path, file_name)

        # Save the article text to the file
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(f"{article_title}\n\n{article_text}")

        print(f"Saved: {file_path}")
        time.sleep(2)
