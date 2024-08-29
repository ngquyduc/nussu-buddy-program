import pandas as pd

match_count = 0
file_path = 'input.xlsx'

# read the excel file
df = pd.read_excel(file_path)

emails = set()
# go through df and count the rows that have 1 in column 'Open 1:2'
for index, row in df.iterrows():
    if row['Open 1:2'] == 1:
        emails.add(row['What is your NUS email address? (e*******@u.nus.edu)'])

print('Number of buddies open to 1:2:', len(emails))


# sort by column 'DateSubmitted'
df = df.sort_values(by='DateSubmitted')

# add new columns
df['Matched buddy email'] = None
df['Matched buddy name'] = None
df['Match with participant 1 email'] = None
df['Match with participant 1 name'] = None
df['Match with participant 2 email'] = None
df['Match with participant 2 name'] = None

# count the total numbers of signups
print('Total number of signups:', len(df))
print("---------------------------------------------------------")

# split the data into 2 parts based on this column What is your year of study?
# group 1 being Masters and group 2 being other categorized as Undergraduate
df_masters = df[df['What is your year of study?'] == 'Masters']
df_undergraduate = df[df['What is your year of study?'] != 'Masters']

# count the number of each group
print('Number of Masters:', len(df_masters))
print('Number of Undergraduate:', len(df_undergraduate))
print("---------------------------------------------------------")

# split the data into two parts based on this column
# (Are you signing up as a Participant (International Student) or Buddy (Existing NUS Student)?)
# column 1 being Participant (International Student) and column 2 being Buddy (Existing NUS Student)
df_masters_participant = df_masters[df_masters['Are you signing up as a Participant (International Student) or Buddy (Existing NUS Student)?']
                                    == 'Participant (International Student)']
df_masters_buddy = df_masters[df_masters['Are you signing up as a Participant (International Student) or Buddy (Existing NUS Student)?']
                              == 'Buddy (Existing NUS Student)']
df_ug_participant = df_undergraduate[df_undergraduate['Are you signing up as a Participant (International Student) or Buddy (Existing NUS Student)?']
                                     == 'Participant (International Student)']
df_ug_buddy = df_undergraduate[df_undergraduate['Are you signing up as a Participant (International Student) or Buddy (Existing NUS Student)?']
                               == 'Buddy (Existing NUS Student)']

# count the numbers of each group
print('Number of Masters Participant:', len(df_masters_participant))
print('Number of Masters Buddy:', len(df_masters_buddy))
print('Number of Undergraduate Participant:', len(df_ug_participant))
print('Number of Undergraduate Buddy:', len(df_ug_buddy))
print("---------------------------------------------------------")

# count the number of master and ug buddies have yes in column open 1:2
master_buddy_open = df_masters_buddy['Open 1:2'] == 1
ug_buddy_open = df_ug_buddy['Open 1:2'] == 1
print('Number of Masters Buddy open to 1:2:',
      len(df_masters_buddy[master_buddy_open]))
print('Number of Undergraduate Buddy open to 1:2:',
      len(df_ug_buddy[ug_buddy_open]))

# create 2 sets of emails for master and ug buddies that are open to 1:2
master_buddy_emails = set(
    df_masters_buddy[master_buddy_open]['What is your NUS email address? (e*******@u.nus.edu)'])
ug_buddy_emails = set(
    df_ug_buddy[ug_buddy_open]['What is your NUS email address? (e*******@u.nus.edu)'])

# count the number of available buddies for master and undergraduate
master_buddy_count = len(df_masters_buddy)
ug_buddy_count = len(df_ug_buddy)

# get the first master_buddy_count number of master participants and ug_buddy_count number of ug participants
df_masters_participant = df_masters_participant.sort_values(by='DateSubmitted')
df_ug_participant = df_ug_participant.sort_values(by='DateSubmitted')
df_masters_participant = df_masters_participant.head(
    master_buddy_count + len(df_masters_buddy[master_buddy_open]))
df_ug_participant = df_ug_participant.head(
    ug_buddy_count + len(df_ug_buddy[ug_buddy_open]))

# split each group to 2 parts based on column ?
# Male and Female
WHAT_GENDER = 'What is your gender?'
MALE = 'Male'
FEMALE = 'Female'
df_masters_participant_male = df_masters_participant[
    df_masters_participant[WHAT_GENDER] == MALE]
df_masters_participant_female = df_masters_participant[
    df_masters_participant[WHAT_GENDER] == FEMALE]
df_masters_buddy_male = df_masters_buddy[df_masters_buddy[WHAT_GENDER] == MALE]
df_masters_buddy_female = df_masters_buddy[df_masters_buddy[WHAT_GENDER] == FEMALE]
df_ug_participant_male = df_ug_participant[df_ug_participant[WHAT_GENDER] == MALE]
df_ug_participant_female = df_ug_participant[df_ug_participant[WHAT_GENDER] == FEMALE]
df_ug_buddy_male = df_ug_buddy[df_ug_buddy[WHAT_GENDER] == MALE]
df_ug_buddy_female = df_ug_buddy[df_ug_buddy[WHAT_GENDER] == FEMALE]

# split the data based on Would you prefer to be matched with the same gender? (Subject to availability)
# group 1 being No preference and group 2 being Yes
GENDER_PREFERENCE = 'Would you prefer to be matched with the same gender? (Subject to availability)'
GENDER_YES = 'Yes'
GENDER_NO_PREF = 'No preference'
df_masters_participant_pref_same_gender = df_masters_participant[
    df_masters_participant[GENDER_PREFERENCE] == GENDER_YES]
df_masters_participant_no_pref = df_masters_participant[
    df_masters_participant[GENDER_PREFERENCE] == GENDER_NO_PREF]
print('Number of Masters Participant with preference:',
      len(df_masters_participant_pref_same_gender))
print('Number of Masters Participant with no preference:',
      len(df_masters_participant_no_pref))
df_masters_buddy_male_pref_male = df_masters_buddy_male[
    df_masters_buddy_male[GENDER_PREFERENCE] == GENDER_YES]
df_masters_buddy_female_pref_female = df_masters_buddy_female[
    df_masters_buddy_female[GENDER_PREFERENCE] == GENDER_YES]

df_ug_participant_pref_same_gender = df_ug_participant[
    df_ug_participant[GENDER_PREFERENCE] == GENDER_YES]
df_ug_participant_no_pref = df_ug_participant[
    df_ug_participant[GENDER_PREFERENCE] == GENDER_NO_PREF]
print('Number of Undergraduate Participant with preference:',
      len(df_ug_participant_pref_same_gender))
print('Number of Undergraduate Participant with no preference:',
      len(df_ug_participant_no_pref))
print('---------------------------------------------------------')
df_ug_buddy_male_pref_male = df_ug_buddy_male[df_ug_buddy_male[GENDER_PREFERENCE] == GENDER_YES]
df_ug_buddy_female_pref_female = df_ug_buddy_female[
    df_ug_buddy_female[GENDER_PREFERENCE] == GENDER_YES]

# interests
INTERESTS = ['Please indicate your top 3 interests from the choices below - Sports',
             'Please indicate your top 3 interests from the choices below - Hiking',
             'Please indicate your top 3 interests from the choices below - Dancing',
             'Please indicate your top 3 interests from the choices below - Food',
             'Please indicate your top 3 interests from the choices below - Museums/Gallery',
             'Please indicate your top 3 interests from the choices below - Photography',
             'Please indicate your top 3 interests from the choices below - Music',
             'Please indicate your top 3 interests from the choices below - Movies',
             'Please indicate your top 3 interests from the choices below - Gym',
             'Please indicate your top 3 interests from the choices below - Art',
             'Please indicate your top 3 interests from the choices below - Gaming',
             'Please indicate your top 3 interests from the choices below - Partying',
             'Please indicate your top 3 interests from the choices below - Others']
SPORTS_YES = 'Sports'
HIKING_YES = 'Hiking'
DANCING_YES = 'Dancing'
FOOD_YES = 'Food'
MUSEUMS_GALLERY_YES = 'Museums/Gallery'
PHOTOGRAPHY_YES = 'Photography'
MUSIC_YES = 'Music'
MOVIES_YES = 'Movies'
GYM_YES = 'Gym'
ART_YES = 'Art'
GAMING_YES = 'Gaming'
PARTYING_YES = 'Partying'
OTHERS_YES = 'Others'

# go through the masters and ugs and create 2 dictionaries of interests
# one for masters and one for ugs
# the key is the 'SubmissionId' and the value is the list of interests
master_buddy_interests = {}
master_participant_interests = {}
ug_buddy_interests = {}
ug_participant_interests = {}
for index, row in df_masters_buddy.iterrows():
    id = row['What is your NUS email address? (e*******@u.nus.edu)']
    interests = []
    for interest in INTERESTS:
        if row[interest] == SPORTS_YES:
            interests.append(SPORTS_YES)
        if row[interest] == HIKING_YES:
            interests.append(HIKING_YES)
        if row[interest] == DANCING_YES:
            interests.append(DANCING_YES)
        if row[interest] == FOOD_YES:
            interests.append(FOOD_YES)
        if row[interest] == MUSEUMS_GALLERY_YES:
            interests.append(MUSEUMS_GALLERY_YES)
        if row[interest] == PHOTOGRAPHY_YES:
            interests.append(PHOTOGRAPHY_YES)
        if row[interest] == MUSIC_YES:
            interests.append(MUSIC_YES)
        if row[interest] == MOVIES_YES:
            interests.append(MOVIES_YES)
        if row[interest] == GYM_YES:
            interests.append(GYM_YES)
        if row[interest] == ART_YES:
            interests.append(ART_YES)
        if row[interest] == GAMING_YES:
            interests.append(GAMING_YES)
        if row[interest] == PARTYING_YES:
            interests.append(PARTYING_YES)
        if row[interest] == OTHERS_YES:
            interests.append(OTHERS_YES)
    master_buddy_interests[id] = interests

for index, row in df_masters_participant.iterrows():
    id = row['What is your NUS email address? (e*******@u.nus.edu)']
    interests = []
    for interest in INTERESTS:
        if row[interest] == SPORTS_YES:
            interests.append(SPORTS_YES)
        if row[interest] == HIKING_YES:
            interests.append(HIKING_YES)
        if row[interest] == DANCING_YES:
            interests.append(DANCING_YES)
        if row[interest] == FOOD_YES:
            interests.append(FOOD_YES)
        if row[interest] == MUSEUMS_GALLERY_YES:
            interests.append(MUSEUMS_GALLERY_YES)
        if row[interest] == PHOTOGRAPHY_YES:
            interests.append(PHOTOGRAPHY_YES)
        if row[interest] == MUSIC_YES:
            interests.append(MUSIC_YES)
        if row[interest] == MOVIES_YES:
            interests.append(MOVIES_YES)
        if row[interest] == GYM_YES:
            interests.append(GYM_YES)
        if row[interest] == ART_YES:
            interests.append(ART_YES)
        if row[interest] == GAMING_YES:
            interests.append(GAMING_YES)
        if row[interest] == PARTYING_YES:
            interests.append(PARTYING_YES)
        if row[interest] == OTHERS_YES:
            interests.append(OTHERS_YES)
    master_participant_interests[id] = interests

for index, row in df_ug_buddy.iterrows():
    id = row['What is your NUS email address? (e*******@u.nus.edu)']
    interests = []
    for interest in INTERESTS:
        if row[interest] == SPORTS_YES:
            interests.append(SPORTS_YES)
        if row[interest] == HIKING_YES:
            interests.append(HIKING_YES)
        if row[interest] == DANCING_YES:
            interests.append(DANCING_YES)
        if row[interest] == FOOD_YES:
            interests.append(FOOD_YES)
        if row[interest] == MUSEUMS_GALLERY_YES:
            interests.append(MUSEUMS_GALLERY_YES)
        if row[interest] == PHOTOGRAPHY_YES:
            interests.append(PHOTOGRAPHY_YES)
        if row[interest] == MUSIC_YES:
            interests.append(MUSIC_YES)
        if row[interest] == MOVIES_YES:
            interests.append(MOVIES_YES)
        if row[interest] == GYM_YES:
            interests.append(GYM_YES)
        if row[interest] == ART_YES:
            interests.append(ART_YES)
        if row[interest] == GAMING_YES:
            interests.append(GAMING_YES)
        if row[interest] == PARTYING_YES:
            interests.append(PARTYING_YES)
        if row[interest] == OTHERS_YES:
            interests.append(OTHERS_YES)
    ug_buddy_interests[id] = interests

for index, row in df_ug_participant.iterrows():
    id = row['What is your NUS email address? (e*******@u.nus.edu)']
    interests = []
    for interest in INTERESTS:
        if row[interest] == SPORTS_YES:
            interests.append(SPORTS_YES)
        if row[interest] == HIKING_YES:
            interests.append(HIKING_YES)
        if row[interest] == DANCING_YES:
            interests.append(DANCING_YES)
        if row[interest] == FOOD_YES:
            interests.append(FOOD_YES)
        if row[interest] == MUSEUMS_GALLERY_YES:
            interests.append(MUSEUMS_GALLERY_YES)
        if row[interest] == PHOTOGRAPHY_YES:
            interests.append(PHOTOGRAPHY_YES)
        if row[interest] == MUSIC_YES:
            interests.append(MUSIC_YES)
        if row[interest] == MOVIES_YES:
            interests.append(MOVIES_YES)
        if row[interest] == GYM_YES:
            interests.append(GYM_YES)
        if row[interest] == ART_YES:
            interests.append(ART_YES)
        if row[interest] == GAMING_YES:
            interests.append(GAMING_YES)
        if row[interest] == PARTYING_YES:
            interests.append(PARTYING_YES)
        if row[interest] == OTHERS_YES:
            interests.append(OTHERS_YES)
    ug_participant_interests[id] = interests


def find_best_interest_matches(pid, from_interests, targeted_interests):
    matches_2 = []
    matches_1 = []
    for buddy_id, interests in targeted_interests.items():
        common_interests = set(
            from_interests[pid]) & set(interests)
        if len(common_interests) == 3:
            # remove from dictionary
            # check if the email of the buddy is in the set of emails
            if pid not in emails:
                del targeted_interests[buddy_id]
            else:
                emails.remove(pid)
            print("Found 3 common interests between", pid, "and", buddy_id)
            return buddy_id
        elif len(common_interests) == 2:
            matches_2.append(buddy_id)
        elif len(common_interests) == 1:
            matches_1.append(buddy_id)
    if len(matches_2) > 0:
        buddy_id = matches_2[0]
        # remove from dictionary
        if pid not in emails:
            del targeted_interests[buddy_id]
        else:
            emails.remove(pid)
        print("Found 2 common interests between", pid, "and", buddy_id)
        return buddy_id
    elif len(matches_1) > 0:
        buddy_id = matches_1[0]
        # remove from dictionary
        if pid not in emails:
            del targeted_interests[buddy_id]
        else:
            emails.remove(pid)
        print("Found 1 common interests between", pid, "and", buddy_id)
        return buddy_id
    else:
        print("No common interests between", pid, "and any buddy")
        return None


# set to count the number of buddies and participants to double check
buddies = set()
participants = set()

# sort all dataframes by 'DateSubmitted'
df_masters_buddy = df_masters_buddy.sort_values(by='DateSubmitted')
df_masters_buddy_male = df_masters_buddy_male.sort_values(by='DateSubmitted')
df_masters_buddy_female = df_masters_buddy_female.sort_values(
    by='DateSubmitted')
df_masters_buddy_male_pref_male = df_masters_buddy_male_pref_male.sort_values(
    by='DateSubmitted')
df_masters_buddy_female_pref_female = df_masters_buddy_female_pref_female.sort_values(
    by='DateSubmitted')
df_masters_participant = df_masters_participant.sort_values(by='DateSubmitted')
df_masters_participant_male = df_masters_participant_male.sort_values(
    by='DateSubmitted')
df_masters_participant_female = df_masters_participant_female.sort_values(
    by='DateSubmitted')
df_masters_participant_pref_same_gender = df_masters_participant_pref_same_gender.sort_values(
    by='DateSubmitted')
df_masters_participant_no_pref = df_masters_participant_no_pref.sort_values(
    by='DateSubmitted')

df_ug_buddy = df_ug_buddy.sort_values(by='DateSubmitted')
df_ug_buddy_male = df_ug_buddy_male.sort_values(by='DateSubmitted')
df_ug_buddy_female = df_ug_buddy_female.sort_values(by='DateSubmitted')
df_ug_buddy_male_pref_male = df_ug_buddy_male_pref_male.sort_values(
    by='DateSubmitted')
df_masters_buddy_female_pref_female = df_masters_buddy_female_pref_female.sort_values(
    by='DateSubmitted')
df_ug_participant = df_ug_participant.sort_values(by='DateSubmitted')
df_ug_participant_male = df_ug_participant_male.sort_values(
    by='DateSubmitted')
df_ug_participant_female = df_ug_participant_female.sort_values(
    by='DateSubmitted')
df_ug_participant_no_pref = df_ug_participant_no_pref.sort_values(
    by='DateSubmitted')
df_ug_participant_pref_same_gender = df_ug_participant_pref_same_gender.sort_values(
    by='DateSubmitted')

# go through the master participants that prefer same gender and match with master buddies that prefer same gender
FULL_NAME = 'What is your full name? (as on your student card)'
NUS_EMAIL = 'What is your NUS email address? (e*******@u.nus.edu)'
for index, row in df_masters_participant_pref_same_gender.iterrows():
    id = row['SubmissionId']
    name = row[FULL_NAME]
    email = row[NUS_EMAIL]
    participants.add(email)
    gender = row['What is your gender?']
    bid = find_best_interest_matches(email, master_participant_interests,
                                     master_buddy_interests)
    if gender == MALE:
        if len(df_masters_buddy_male_pref_male) > 0:
            if bid is not None:
                mask = df_masters_buddy_male_pref_male['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_masters_buddy_male_pref_male[mask].iloc[0]
            else:
                buddy = df_masters_buddy_male_pref_male.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            # remove the matched buddy from the dataframes
            if buddy_email not in emails:
                f1 = df_masters_buddy_male_pref_male[NUS_EMAIL] != buddy_email
                df_masters_buddy_male_pref_male = df_masters_buddy_male_pref_male[f1]
                f2 = df_masters_buddy_male[NUS_EMAIL] != buddy_email
                df_masters_buddy_male = df_masters_buddy_male[f2]
                f3 = df_masters_buddy[NUS_EMAIL] != buddy_email
                df_masters_buddy = df_masters_buddy[f3]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        elif len(df_masters_buddy_male) > 0:
            if bid is not None:
                mask = df_masters_buddy_male['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_masters_buddy_male[mask].iloc[0]
            else:
                buddy = df_masters_buddy_male.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_masters_buddy_male[NUS_EMAIL] != buddy_email
                df_masters_buddy_male = df_masters_buddy_male[f1]
                f2 = df_masters_buddy[NUS_EMAIL] != buddy_email
                df_masters_buddy = df_masters_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        else:
            if bid is not None:
                mask = df_masters_buddy_female['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_masters_buddy_female[mask].iloc[0]
            else:
                buddy = df_masters_buddy_female.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_masters_buddy_female[NUS_EMAIL] != buddy_email
                df_masters_buddy_female = df_masters_buddy_female[f1]
                f2 = df_masters_buddy[NUS_EMAIL] != buddy_email
                df_masters_buddy = df_masters_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)
    else:
        if len(df_masters_buddy_female_pref_female) > 0:
            if bid is not None:
                mask = df_masters_buddy_female_pref_female['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_masters_buddy_female_pref_female[mask].iloc[0]
            else:
                buddy = df_masters_buddy_female_pref_female.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_masters_buddy_female_pref_female[NUS_EMAIL] != buddy_email
                df_masters_buddy_female_pref_female = df_masters_buddy_female_pref_female[f1]
                f2 = df_masters_buddy_female[NUS_EMAIL] != buddy_email
                df_masters_buddy_female = df_masters_buddy_female[f2]
                f3 = df_masters_buddy[NUS_EMAIL] != buddy_email
                df_masters_buddy = df_masters_buddy[f3]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        elif len(df_masters_buddy_female) > 0:
            if bid is not None:
                mask = df_masters_buddy_female['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_masters_buddy_female[mask].iloc[0]
            else:
                buddy = df_masters_buddy_female.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_masters_buddy_female[NUS_EMAIL] != buddy_email
                df_masters_buddy_female = df_masters_buddy_female[f1]
                f2 = df_masters_buddy[NUS_EMAIL] != buddy_email
                df_masters_buddy = df_masters_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        else:
            if bid is not None:
                mask = df_masters_buddy_male['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_masters_buddy_male[mask].iloc[0]
            else:
                buddy = df_masters_buddy_male.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_masters_buddy_male[NUS_EMAIL] != buddy_email
                df_masters_buddy_male = df_masters_buddy_male[f1]
                f2 = df_masters_buddy[NUS_EMAIL] != buddy_email
                df_masters_buddy = df_masters_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

print('---------------------------------------------------------')

# go through the master participants that prefer no preference and match with the rest
for index, row in df_masters_participant_no_pref.iterrows():
    id = row['SubmissionId']
    name = row[FULL_NAME]
    email = row[NUS_EMAIL]
    participants.add(email)
    bid = find_best_interest_matches(
        email, master_participant_interests, master_buddy_interests)
    if bid is not None:
        mask = df_masters_buddy['SubmissionId'] == bid
    if bid is not None and mask is not None and mask.any():
        buddy = df_masters_buddy[mask].iloc[0]
    else:
        buddy = df_masters_buddy.iloc[0]
    buddy_id = buddy['SubmissionId']
    buddy_name = buddy[FULL_NAME]
    buddy_email = buddy[NUS_EMAIL]
    df.loc[df['SubmissionId'] == id, 'Match with buddy email'] = buddy_email
    df.loc[df['SubmissionId'] == id, 'Match with buddy name'] = buddy_name
    if buddy_email not in emails:
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 1 email'] = email
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 1 name'] = name
    else:
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 2 email'] = email
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 2 name'] = name
    if buddy_email not in emails:
        f = df_masters_buddy[NUS_EMAIL] != buddy_email
        df_masters_buddy = df_masters_buddy[f]
    else:
        emails.remove(buddy_email)
    match_count += 1
    if buddy_email in buddies:
        print('DUPLICATE:', buddy_email)
    buddies.add(buddy_email)
    print(match_count, 'Matched', name, 'with', buddy_name)

print('---------------------------------------------------------')
# go through the ug participants that prefer same gender
for index, row in df_ug_participant_pref_same_gender.iterrows():
    id = row['SubmissionId']
    name = row[FULL_NAME]
    email = row[NUS_EMAIL]
    participants.add(email)
    gender = row['What is your gender?']
    bid = find_best_interest_matches(
        email, ug_participant_interests, ug_buddy_interests)
    if gender == MALE:
        if len(df_ug_buddy_male_pref_male) > 0:
            if bid is not None:
                mask = df_ug_buddy_male_pref_male['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_ug_buddy_male_pref_male[mask].iloc[0]
            else:
                buddy = df_ug_buddy_male_pref_male.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_ug_buddy_male_pref_male[NUS_EMAIL] != buddy_email
                df_ug_buddy_male_pref_male = df_ug_buddy_male_pref_male[f1]
                f2 = df_ug_buddy_male[NUS_EMAIL] != buddy_email
                df_ug_buddy_male = df_ug_buddy_male[f2]
                f3 = df_ug_buddy[NUS_EMAIL] != buddy_email
                df_ug_buddy = df_ug_buddy[f3]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        elif len(df_ug_buddy_male) > 0:
            if bid is not None:
                mask = df_ug_buddy_male['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_ug_buddy_male[mask].iloc[0]
            else:
                buddy = df_ug_buddy_male.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_ug_buddy_male[NUS_EMAIL] != buddy_email
                df_ug_buddy_male = df_ug_buddy_male[f1]
                f2 = df_ug_buddy[NUS_EMAIL] != buddy_email
                df_ug_buddy = df_ug_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        else:
            if bid is not None:
                mask = df_ug_buddy_female['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_ug_buddy_female[mask].iloc[0]
            else:
                buddy = df_ug_buddy_female.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_ug_buddy_female[NUS_EMAIL] != buddy_email
                df_ug_buddy_female = df_ug_buddy_female[f1]
                f2 = df_ug_buddy[NUS_EMAIL] != buddy_email
                df_ug_buddy = df_ug_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)
    else:
        if len(df_ug_buddy_female_pref_female) > 0:
            if bid is not None:
                mask = df_ug_buddy_female_pref_female['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_ug_buddy_female_pref_female[mask].iloc[0]
            else:
                buddy = df_ug_buddy_female_pref_female.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_ug_buddy_female_pref_female[NUS_EMAIL] != buddy_email
                df_ug_buddy_female_pref_female = df_ug_buddy_female_pref_female[f1]
                f2 = df_ug_buddy_female[NUS_EMAIL] != buddy_email
                df_ug_buddy_female = df_ug_buddy_female[f2]
                f3 = df_ug_buddy[NUS_EMAIL] != buddy_email
                df_ug_buddy = df_ug_buddy[f3]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        elif len(df_ug_buddy_female) > 0:
            if bid is not None:
                mask = df_ug_buddy_female['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_ug_buddy_female[mask].iloc[0]
            else:
                buddy = df_ug_buddy_female.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_ug_buddy_female[NUS_EMAIL] != buddy_email
                df_ug_buddy_female = df_ug_buddy_female[f1]
                f2 = df_ug_buddy[NUS_EMAIL] != buddy_email
                df_ug_buddy = df_ug_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

        else:
            if bid is not None:
                mask = df_ug_buddy_male['SubmissionId'] == bid
            if bid is not None and mask is not None and mask.any():
                buddy = df_ug_buddy_male[mask].iloc[0]
            else:
                buddy = df_ug_buddy_male.iloc[0]
            buddy_id = buddy['SubmissionId']
            buddy_name = buddy[FULL_NAME]
            buddy_email = buddy[NUS_EMAIL]
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy email'] = buddy_email
            df.loc[df['SubmissionId'] == id,
                   'Match with buddy name'] = buddy_name
            if buddy_email not in emails:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 1 name'] = name
            else:
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 email'] = email
                df.loc[df['SubmissionId'] == buddy_id,
                       'Match with participant 2 name'] = name
            if buddy_email not in emails:
                f1 = df_ug_buddy_male[NUS_EMAIL] != buddy_email
                df_ug_buddy_male = df_ug_buddy_male[f1]
                f2 = df_ug_buddy[NUS_EMAIL] != buddy_email
                df_ug_buddy = df_ug_buddy[f2]
            else:
                emails.remove(buddy_email)
            match_count += 1
            if buddy_email in buddies:
                print('DUPLICATE:', buddy_email)
            buddies.add(buddy_email)
            print(match_count, 'Matched', name, 'with', buddy_name)

print('---------------------------------------------------------')

for index, row in df_ug_participant_no_pref.iterrows():
    id = row['SubmissionId']
    name = row[FULL_NAME]
    email = row[NUS_EMAIL]
    participants.add(email)
    bid = find_best_interest_matches(
        email, ug_participant_interests, ug_buddy_interests)
    if bid is not None:
        mask = df_ug_buddy['SubmissionId'] == bid
    if bid is not None and mask is not None and mask.any():
        buddy = df_ug_buddy[mask].iloc[0]
    else:
        buddy = df_ug_buddy.iloc[0]
    buddy_id = buddy['SubmissionId']
    buddy_name = buddy[FULL_NAME]
    buddy_email = buddy[NUS_EMAIL]
    df.loc[df['SubmissionId'] == id, 'Match with buddy email'] = buddy_email
    df.loc[df['SubmissionId'] == id, 'Match with buddy name'] = buddy_name
    if buddy_email not in emails:
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 1 email'] = email
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 1 name'] = name
    else:
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 2 email'] = email
        df.loc[df['SubmissionId'] == buddy_id,
               'Match with participant 2 name'] = name
    if buddy_email not in emails:
        f = df_ug_buddy[NUS_EMAIL] != buddy_email
        df_ug_buddy = df_ug_buddy[f]
    else:
        emails.remove(buddy_email)
    match_count += 1
    if buddy_email in buddies:
        print('DUPLICATE:', buddy_email)
    buddies.add(buddy_email)
    print(match_count, 'Matched', name, 'with', buddy_name)


print('TOTAL MATCHED:', match_count, '/',
      len(df_masters_participant) + len(df_ug_participant))
print('TOTAL PARTICIPANTS:', len(participants))
print('TOTAL BUDDIES:', len(buddies))

# write back to the file
df = df.sort_values(by='DateSubmitted')
df.to_excel('output.xlsx', index=False)
