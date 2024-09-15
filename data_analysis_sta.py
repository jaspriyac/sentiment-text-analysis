#data analysis
import os
import re
import pandas as pd
import xlsxwriter 
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords

# Set the path to your desktop
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

# Set the path to the directory containing text files
files_directory = os.path.join(desktop_path, 'blackassign_articles')

# Set the path to the stopwords directory
stopwords_directory = os.path.join(desktop_path, 'StopWords')

# Set the path to the master dictionary directory
master_dict_directory = os.path.join(desktop_path, 'MasterDictionary')

# Function to read file
def read_file(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
        return file.read()
    
# Function to save output to excel file
def save_to_excel(output_data, excel_path):
    df = pd.DataFrame(output_data)
    df.to_excel(excel_path, index=False)
    
# Function to clean and tokenize text
def clean_text(text, stop_words):
    # Tokenize the text
    tokens = word_tokenize(text.lower())

    # Remove stopwords
    tokens = [word for word in tokens if word.isalnum() and word not in stop_words]

    return tokens

# Function to load stopwords from various lists
def get_stopwords(stopwords_directory):
    stop_words = set()
    stopwords_files = ['StopWords_Auditor.txt', 'StopWords_Currencies.txt', 'StopWords_DatesandNumbers.txt',
                        'StopWords_Generic.txt', 'StopWords_GenericLong.txt', 'StopWords_Geographic.txt',
                        'StopWords_Names.txt']

    for filename in stopwords_files:
        filepath = os.path.join(stopwords_directory, filename)
        stop_words.update(set(word_tokenize(read_file(filepath).lower())))

    return stop_words

# Function to load positive and negative list
def get_master_dictionary(master_dict_directory):
    master_dict = {'positive': set(), 'negative': set()}

    for sentiment in ['positive', 'negative']:
        filename = f'{sentiment}-words.txt'
        filepath = os.path.join(master_dict_directory, filename)
        master_dict[sentiment].update(set(word_tokenize(read_file(filepath).lower())))

    return master_dict

# Function to calculate text analysis metrics
def calculate_scores(tokens, positive_dict, negative_dict):
    try:
        # Calculate positive score, negative score
        positive_score = sum(1 for word in tokens if word in positive_dict)
        negative_score = (-sum(1 for word in tokens if word in negative_dict)*-1)

        # Calculate word count
        total_words = len(tokens)
    
        # Calculate average sentence length
        avg_sentence_length = total_words / len(sent_tokenize(text))
    
        # Calculate complex words, complex words count and percentage of complex words
        complex_words = [word for word in tokens if syllable_count(word) > 2]
        complex_word_count = len(complex_words)
        percentage_complex_words = complex_word_count / total_words 
    
        # Calculate fog index
        fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
        # Calculate average number of words per sentence
        avg_num_words_per_sentence = total_words / len(sent_tokenize(text))
    
        # Calculate syllable count per word
        syllable_per_word = sum(syllable_count(word) for word in tokens) / total_words
    
        # Calculate personal pronoun count
        personal_pronouns = len(re.findall(r'\b(i|we|my|ours|us)\b', text, flags=re.IGNORECASE))

        # Calculate average word length
        avg_word_length = sum(len(word) for word in tokens) / total_words

        # Calculate polarity score and subjectivity score
        subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
        polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)

        return {
            'positive_score': positive_score,
            'negative_score': negative_score,
            'polarity_score': polarity_score,
            'subjectivity_score': subjectivity_score,
            'avg_sentence_length': avg_sentence_length,
            'percentage_complex_words': percentage_complex_words,
            'fog_index': fog_index,
            'avg_num_words_per_sentence': avg_num_words_per_sentence,
            'complex_word_count': complex_word_count,
            'total_words': total_words,
            'syllable_per_word': syllable_per_word,
            'personal_pronouns': personal_pronouns,
            'avg_word_length': avg_word_length,
        }
    except Exception as e:
        print(f"Error calculating scores: {e}")
        return None

# Function to count syllables in a word
def syllable_count(word):
    vowels = "aeiouy"
    exceptions = ["es", "ed"]
    count = 0
    word = word.lower()
    
    # Check for exceptions
    for exception in exceptions:
        if word.endswith(exception):
            return count
    
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e") and not word.endswith("le"):
        count -= 1
    if count == 0:
        count += 1
    return count


# Get stopwords
stop_words = get_stopwords(stopwords_directory)

# Get master dictionary
master_dict = get_master_dictionary(master_dict_directory)

# Initialize scores
positive_score = 0
negative_score = 0

# Initialize a DataFrame to store the results
columns = ['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE',
           'SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS',
           'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT',
           'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']
output_df = pd.DataFrame(columns=columns)


# Input 100 URL IDs and URLs
url_data = [
    ('blackassign0001', 'https://insights.blackcoffer.com/rising-it-cities-and-its-impact-on-the-economy-environment-infrastructure-and-city-life-by-the-year-2040-2/'),
    ('blackassign0002', 'https://insights.blackcoffer.com/rising-it-cities-and-their-impact-on-the-economy-environment-infrastructure-and-city-life-in-future/'),
    ('blackassign0003', 'https://insights.blackcoffer.com/internet-demands-evolution-communication-impact-and-2035s-alternative-pathways/'),
    ('blackassign0004', 'https://insights.blackcoffer.com/rise-of-cybercrime-and-its-effect-in-upcoming-future/'),
    ('blackassign0005', 'https://insights.blackcoffer.com/ott-platform-and-its-impact-on-the-entertainment-industry-in-future/'),
    ('blackassign0006', 'https://insights.blackcoffer.com/the-rise-of-the-ott-platform-and-its-impact-on-the-entertainment-industry-by-2040/'),
    ('blackassign0007', 'https://insights.blackcoffer.com/rise-of-cyber-crime-and-its-effects/'),
    ('blackassign0008', 'https://insights.blackcoffer.com/rise-of-internet-demand-and-its-impact-on-communications-and-alternatives-by-the-year-2035-2/'),
    ('blackassign0009', 'https://insights.blackcoffer.com/rise-of-cybercrime-and-its-effect-by-the-year-2040-2/'),
    ('blackassign0010', 'https://insights.blackcoffer.com/rise-of-cybercrime-and-its-effect-by-the-year-2040/'),
    ('blackassign0011', 'https://insights.blackcoffer.com/rise-of-internet-demand-and-its-impact-on-communications-and-alternatives-by-the-year-2035/'),
    ('blackassign0012', 'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-3-2/'),
    ('blackassign0013', 'https://insights.blackcoffer.com/rise-of-e-health-and-its-impact-on-humans-by-the-year-2030/'),
    ('blackassign0014', 'https://insights.blackcoffer.com/rise-of-e-health-and-its-imapct-on-humans-by-the-year-2030-2/'),
    ('blackassign0015', 'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-2/'),
    ('blackassign0016', 'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-2-2/'),
    ('blackassign0017', 'https://insights.blackcoffer.com/rise-of-chatbots-and-its-impact-on-customer-support-by-the-year-2040/'),
    ('blackassign0018', 'https://insights.blackcoffer.com/rise-of-e-health-and-its-imapct-on-humans-by-the-year-2030/'),
    ('blackassign0019', 'https://insights.blackcoffer.com/how-does-marketing-influence-businesses-and-consumers/'),
    ('blackassign0020', 'https://insights.blackcoffer.com/how-advertisement-increase-your-market-value/'),
    ('blackassign0021', 'https://insights.blackcoffer.com/negative-effects-of-marketing-on-society/'),
    ('blackassign0022', 'https://insights.blackcoffer.com/how-advertisement-marketing-affects-business/'),
    ('blackassign0023', 'https://insights.blackcoffer.com/rising-it-cities-will-impact-the-economy-environment-infrastructure-and-city-life-by-the-year-2035/'),
    ('blackassign0024', 'https://insights.blackcoffer.com/rise-of-ott-platform-and-its-impact-on-entertainment-industry-by-the-year-2030/'),
    ('blackassign0025', 'https://insights.blackcoffer.com/rise-of-electric-vehicles-and-its-impact-on-livelihood-by-2040/'),
    ('blackassign0026', 'https://insights.blackcoffer.com/rise-of-electric-vehicle-and-its-impact-on-livelihood-by-the-year-2040/'),
    ('blackassign0027', 'https://insights.blackcoffer.com/oil-prices-by-the-year-2040-and-how-it-will-impact-the-world-economy/'),
    ('blackassign0028', 'https://insights.blackcoffer.com/an-outlook-of-healthcare-by-the-year-2040-and-how-it-will-impact-human-lives/'),
    ('blackassign0029', 'https://insights.blackcoffer.com/ai-in-healthcare-to-improve-patient-outcomes/'),
    ('blackassign0030', 'https://insights.blackcoffer.com/what-if-the-creation-is-taking-over-the-creator/'),
    ('blackassign0031', 'https://insights.blackcoffer.com/what-jobs-will-robots-take-from-humans-in-the-future/'),
    ('blackassign0032', 'https://insights.blackcoffer.com/will-machine-replace-the-human-in-the-future-of-work/'),
    ('blackassign0033', 'https://insights.blackcoffer.com/will-ai-replace-us-or-work-with-us/'),
    ('blackassign0034', 'https://insights.blackcoffer.com/man-and-machines-together-machines-are-more-diligent-than-humans-blackcoffe/'),
    ('blackassign0035', 'https://insights.blackcoffer.com/in-future-or-in-upcoming-years-humans-and-machines-are-going-to-work-together-in-every-field-of-work/'),
    ('blackassign0036', 'https://insights.blackcoffer.com/how-neural-networks-can-be-applied-in-various-areas-in-the-future/'),
    ('blackassign0037', 'https://insights.blackcoffer.com/how-machine-learning-will-affect-your-business/'),
    ('blackassign0038', 'https://insights.blackcoffer.com/deep-learning-impact-on-areas-of-e-learning/'),
    ('blackassign0039', 'https://insights.blackcoffer.com/how-to-protect-future-data-and-its-privacy-blackcoffer/'),
    ('blackassign0040', 'https://insights.blackcoffer.com/how-machines-ai-automations-and-robo-human-are-effective-in-finance-and-banking/'),
    ('blackassign0041', 'https://insights.blackcoffer.com/ai-human-robotics-machine-future-planet-blackcoffer-thinking-jobs-workplace/'),
    ('blackassign0042', 'https://insights.blackcoffer.com/how-ai-will-change-the-world-blackcoffer/'),
    ('blackassign0043', 'https://insights.blackcoffer.com/future-of-work-how-ai-has-entered-the-workplace/'),
    ('blackassign0044', 'https://insights.blackcoffer.com/ai-tool-alexa-google-assistant-finance-banking-tool-future/'),
    ('blackassign0045', 'https://insights.blackcoffer.com/ai-healthcare-revolution-ml-technology-algorithm-google-analytics-industrialrevolution/'),
    ('blackassign0046', 'https://insights.blackcoffer.com/all-you-need-to-know-about-online-marketing/'),
    ('blackassign0047', 'https://insights.blackcoffer.com/evolution-of-advertising-industry/'),
    ('blackassign0048', 'https://insights.blackcoffer.com/how-data-analytics-can-help-your-business-respond-to-the-impact-of-covid-19/'),
    ('blackassign0049', 'https://insights.blackcoffer.com/covid-19-environmental-impact-for-the-future/'),
    ('blackassign0050', 'https://insights.blackcoffer.com/environmental-impact-of-the-covid-19-pandemic-lesson-for-the-future/'),
    ('blackassign0051', 'https://insights.blackcoffer.com/how-data-analytics-and-ai-are-used-to-halt-the-covid-19-pandemic/'),
    ('blackassign0052', 'https://insights.blackcoffer.com/difference-between-artificial-intelligence-machine-learning-statistics-and-data-mining/'),
    ('blackassign0053', 'https://insights.blackcoffer.com/how-python-became-the-first-choice-for-data-science/'),
    ('blackassign0054', 'https://insights.blackcoffer.com/how-google-fit-measure-heart-and-respiratory-rates-using-a-phone/'),
    ('blackassign0055', 'https://insights.blackcoffer.com/what-is-the-future-of-mobile-apps/'),
    ('blackassign0056', 'https://insights.blackcoffer.com/impact-of-ai-in-health-and-medicine/'),
    ('blackassign0057', 'https://insights.blackcoffer.com/telemedicine-what-patients-like-and-dislike-about-it/'),
    ('blackassign0058', 'https://insights.blackcoffer.com/how-we-forecast-future-technologies/'),
    ('blackassign0059', 'https://insights.blackcoffer.com/can-robots-tackle-late-life-loneliness/'),
    ('blackassign0060', 'https://insights.blackcoffer.com/embedding-care-robots-into-society-socio-technical-considerations/'),
    ('blackassign0061', 'https://insights.blackcoffer.com/management-challenges-for-future-digitalization-of-healthcare-services/'),
    ('blackassign0062', 'https://insights.blackcoffer.com/are-we-any-closer-to-preventing-a-nuclear-holocaust/'),
    ('blackassign0063', 'https://insights.blackcoffer.com/will-technology-eliminate-the-need-for-animal-testing-in-drug-development/'),
    ('blackassign0064', 'https://insights.blackcoffer.com/will-we-ever-understand-the-nature-of-consciousness/'),
    ('blackassign0065', 'https://insights.blackcoffer.com/will-we-ever-colonize-outer-space/'),
    ('blackassign0066', 'https://insights.blackcoffer.com/what-is-the-chance-homo-sapiens-will-survive-for-the-next-500-years/'),
    ('blackassign0067', 'https://insights.blackcoffer.com/why-does-your-business-need-a-chatbot/'),
    ('blackassign0068', 'https://insights.blackcoffer.com/how-you-lead-a-project-or-a-team-without-any-technical-expertise/'),
    ('blackassign0069', 'https://insights.blackcoffer.com/can-you-be-great-leader-without-technical-expertise/'),
    ('blackassign0070', 'https://insights.blackcoffer.com/how-does-artificial-intelligence-affect-the-environment/'),
    ('blackassign0071', 'https://insights.blackcoffer.com/how-to-overcome-your-fear-of-making-mistakes-2/'),
    ('blackassign0072', 'https://insights.blackcoffer.com/is-perfection-the-greatest-enemy-of-productivity/'),
    ('blackassign0073', 'https://insights.blackcoffer.com/global-financial-crisis-2008-causes-effects-and-its-solution/'),
    ('blackassign0074', 'https://insights.blackcoffer.com/gender-diversity-and-equality-in-the-tech-industry/'),
    ('blackassign0075', 'https://insights.blackcoffer.com/how-to-overcome-your-fear-of-making-mistakes/'),
    ('blackassign0076', 'https://insights.blackcoffer.com/how-small-business-can-survive-the-coronavirus-crisis/'),
    ('blackassign0077', 'https://insights.blackcoffer.com/impacts-of-covid-19-on-vegetable-vendors-and-food-stalls/'),
    ('blackassign0078', 'https://insights.blackcoffer.com/impacts-of-covid-19-on-vegetable-vendors/'),
    ('blackassign0079', 'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-tourism-aviation-industries/'),
    ('blackassign0080', 'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-sports-events-around-the-world/'),
    ('blackassign0081', 'https://insights.blackcoffer.com/changing-landscape-and-emerging-trends-in-the-indian-it-ites-industry/'),
    ('blackassign0082', 'https://insights.blackcoffer.com/online-gaming-adolescent-online-gaming-effects-demotivated-depression-musculoskeletal-and-psychosomatic-symptoms/'),
    ('blackassign0083', 'https://insights.blackcoffer.com/human-rights-outlook/'),
    ('blackassign0084', 'https://insights.blackcoffer.com/how-voice-search-makes-your-business-a-successful-business/'),
    ('blackassign0085', 'https://insights.blackcoffer.com/how-the-covid-19-crisis-is-redefining-jobs-and-services/'),
    ('blackassign0086', 'https://insights.blackcoffer.com/how-to-increase-social-media-engagement-for-marketers/'),
    ('blackassign0087', 'https://insights.blackcoffer.com/impacts-of-covid-19-on-streets-sides-food-stalls/'),
    ('blackassign0088', 'https://insights.blackcoffer.com/coronavirus-impact-on-energy-markets-2/'),
    ('blackassign0089', 'https://insights.blackcoffer.com/coronavirus-impact-on-the-hospitality-industry-5/'),
    ('blackassign0090', 'https://insights.blackcoffer.com/lessons-from-the-past-some-key-learnings-relevant-to-the-coronavirus-crisis-4/'),
    ('blackassign0091', 'https://insights.blackcoffer.com/estimating-the-impact-of-covid-19-on-the-world-of-work-2/'),
    ('blackassign0092', 'https://insights.blackcoffer.com/estimating-the-impact-of-covid-19-on-the-world-of-work-3/'),
    ('blackassign0093', 'https://insights.blackcoffer.com/travel-and-tourism-outlook/'),
    ('blackassign0094', 'https://insights.blackcoffer.com/gaming-disorder-and-effects-of-gaming-on-health/'),
    ('blackassign0095', 'https://insights.blackcoffer.com/what-is-the-repercussion-of-the-environment-due-to-the-covid-19-pandemic-situation/'),
    ('blackassign0096', 'https://insights.blackcoffer.com/what-is-the-repercussion-of-the-environment-due-to-the-covid-19-pandemic-situation-2/'),
    ('blackassign0097', 'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-office-space-and-co-working-industries/'),
    ('blackassign0098', 'https://insights.blackcoffer.com/contribution-of-handicrafts-visual-arts-literature-in-the-indian-economy/'),
    ('blackassign0099', 'https://insights.blackcoffer.com/how-covid-19-is-impacting-payment-preferences/'),
    ('blackassign0100', 'https://insights.blackcoffer.com/how-will-covid-19-affect-the-world-of-work-2/')
]



# Iterate through text files
for url_id, url in url_data:
    file_name = f"{url_id}.txt"
    file_path = os.path.join(files_directory, file_name)

    # Check if the file exists
    if not os.path.isfile(file_path):
        print(f"File not found: {file_path}")
        continue

    # Read file
    text = read_file(file_path)

    # Clean text
    tokens = clean_text(text, stop_words)

    # Calculate scores
    scores = calculate_scores(tokens, master_dict['positive'], master_dict['negative'])

    # Assign 0 values for specific URL_IDs
    if url_id in ['blackassign0036', 'blackassign0049']:
        scores = {key: 0 for key in scores}

    # Append scores to DataFrame
    output_df = output_df.append({
        'URL_ID': url_id,
        'URL': url,
        'POSITIVE SCORE': scores.get('positive_score', 0),
        'NEGATIVE SCORE': scores.get('negative_score', 0),
        'POLARITY SCORE': scores.get('polarity_score', 0),
        'SUBJECTIVITY SCORE': scores.get('subjectivity_score', 0),
        'AVG SENTENCE LENGTH': scores.get('avg_sentence_length', 0),
        'PERCENTAGE OF COMPLEX WORDS': scores.get('percentage_complex_words', 0),
        'FOG INDEX': scores.get('fog_index', 0),
        'AVG NUMBER OF WORDS PER SENTENCE': scores.get('avg_num_words_per_sentence', 0),
        'COMPLEX WORD COUNT': scores.get('complex_word_count', 0),
        'WORD COUNT': scores.get('total_words', 0),
        'SYLLABLE PER WORD': scores.get('syllable_per_word', 0),
        'PERSONAL PRONOUNS': scores.get('personal_pronouns', 0),
        'AVG WORD LENGTH': scores.get('avg_word_length', 0),
    }, ignore_index=True)
# Save the DataFrame to an Excel file with hyperlinks and left-aligned column names
output_file_path = os.path.join(desktop_path, 'Output Data Structure.xlsx')
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # Append scores to DataFrame
    new_entries = [
        {
            'URL_ID': 'blackassign0036',
            'URL': 'https://insights.blackcoffer.com/how-neural-networks-can-be-applied-in-various-areas-in-the-future/',
            'POSITIVE SCORE': 0,
            'NEGATIVE SCORE': 0,
            'POLARITY SCORE': 0,
            'SUBJECTIVITY SCORE': 0,
            'AVG SENTENCE LENGTH': 0,
            'PERCENTAGE OF COMPLEX WORDS': 0,
            'FOG INDEX': 0,
            'AVG NUMBER OF WORDS PER SENTENCE': 0,
            'COMPLEX WORD COUNT': 0,
            'WORD COUNT': 0,
            'SYLLABLE PER WORD': 0,
            'PERSONAL PRONOUNS': 0,
            'AVG WORD LENGTH': 0,
        },
        {
            'URL_ID': 'blackassign0049',
            'URL': 'https://insights.blackcoffer.com/covid-19-environmental-impact-for-the-future/',
            'POSITIVE SCORE': 0,
            'NEGATIVE SCORE': 0,
            'POLARITY SCORE': 0,
            'SUBJECTIVITY SCORE': 0,
            'AVG SENTENCE LENGTH': 0,
            'PERCENTAGE OF COMPLEX WORDS': 0,
            'FOG INDEX': 0,
            'AVG NUMBER OF WORDS PER SENTENCE': 0,
            'COMPLEX WORD COUNT': 0,
            'WORD COUNT': 0,
            'SYLLABLE PER WORD': 0,
            'PERSONAL PRONOUNS': 0,
            'AVG WORD LENGTH': 0,
        }
    ]

    new_df = pd.DataFrame(new_entries)
    output_df = pd.concat([output_df, new_df]).sort_values(by='URL_ID')

    output_df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Apply the URL format (blue color) to the 'URL' column
    url_format = workbook.add_format({'color': 'blue', 'underline': 1, 'align': 'left'})
    worksheet.set_column('B:B', None, url_format)

print(f"Output saved to {output_file_path}")
