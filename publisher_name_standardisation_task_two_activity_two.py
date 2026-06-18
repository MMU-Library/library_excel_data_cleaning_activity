#import pandas package, fuzzywuzzy for string matching, and pip to install packages
import pip
import ensurepip
ensurepip
import re
from gettext import install
import rapidfuzz
from rapidfuzz import fuzz, process
import pandas as pd
from collections import Counter
 


#-------------------import file and assign it to a variable as a data frame ---------------------------------

df = metadata_extract = pd.read_excel(r'C:\Users\46071956\Downloads\library_excel_data_cleaning_scripts\Outputs\metadata_extract_20260127_filtered_task_two_activity_one.xlsx')



#----------------store the values within the "publisher" column as a variable--------------------------------
publisher_col="publisher"


#----------------define a cleaning function that checks for the following:-----------------------------------
# Identify None/NaN values 
# Covert values to text and lowercase
# Remove punctuation
# Make '&' and 'And' equivalent
# Remove extra spaces 
# Remove corporate markers such as 'ltd', 'plc, 'gmbh'
#----------------finally, return back the cleaned publisher names-------------------------------------------


def clean_publisher (text):
    if pd.isna(text):
        return None
    text =str(text).lower().strip()
    text =re.sub(r"[^\w\s&]"," ",text)
    text =text.replace("&", " and ")
    text =re.sub(r"\s+"," ",text).strip()
    text =re.sub(r"\b(ltd|limited|gmbh|corp|inc)\b"," ",text)
    text =re.sub(r"\s+"," ",text).strip()
    return text


df["publisher_cleaned"]=df[publisher_col].apply(clean_publisher)

# ---------- 3.Unique names + frequency-----------------

unique_pubs = df["publisher_cleaned"].dropna().tolist()
freq_counter = Counter(unique_pubs)
unique_pubs = list(freq_counter.keys())
n = len(unique_pubs)
indices = list(range(n))

#-------------- 4. Configuration--------------------

SIMILARITY_THRESHOLD_AUTO = 85
SIMILARITY_THRESHOLD_REVIEW = 65
FREQ_RATIO_FOR_AUTO = 5

STOPWORDS = {
    "the", "of", "and", "for", "on", "at", "to", "in", "with", "without",
    "publishing", "press", "group", "inc", "ltd", "limited", "corporation",
    "university", "college", "school", "institute", "journal", "international",
    "national", "centre", "center", "association", "society"
}

#
#------- 5. Token preprocessing--------------------------

def preprocess_name(name):
    if pd.isna(name):
        return []
    s = str(name).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    tokens = s.split()
    tokens = [t for t in tokens if t not in STOPWORDS]
    return tokens

def get_group_key(name):
    tokens = preprocess_name(name)
    if tokens:
        return tokens[0]
    return str(name)[:3].upper()


# ------------------6. Pairwise matching--------------------

auto_edges = []
review_pairs = []

for a in range(n):
    for b in range(a + 1, n):
        idx_i, idx_j = indices[a], indices[b]
        name_i, name_j = unique_pubs[idx_i], unique_pubs[idx_j]

        tokens_i = preprocess_name(name_i)
        tokens_j = preprocess_name(name_j)

        sim = fuzz.token_sort_ratio(" ".join(tokens_i), " ".join(tokens_j))

        first_token_match = bool(tokens_i and tokens_j and tokens_i[0] == tokens_j[0])

        freq_i = freq_counter[name_i]
        freq_j = freq_counter[name_j]
        freq_ratio = max(freq_i, freq_j) / min(freq_i, freq_j) if min(freq_i, freq_j) > 0 else 0

        if sim >= SIMILARITY_THRESHOLD_AUTO:
            auto_edges.append((idx_i, idx_j, sim))

        elif sim >= SIMILARITY_THRESHOLD_REVIEW:
            if (sim >= 70 and freq_ratio >= FREQ_RATIO_FOR_AUTO) or (first_token_match and sim >= 75):
                auto_edges.append((idx_i, idx_j, sim))
            else:
                review_pairs.append((idx_i, idx_j, sim, name_i, name_j, freq_i, freq_j, first_token_match))



#--------------------store each unique publisher name in a list, dropping any NaN values and duplicates-----------------


unique_publisher_names=df["publisher_cleaned"].dropna().drop_duplicates().tolist()
print("Unique publisher names:", unique_publisher_names)
print("Number of unique publisher names:", len(unique_publisher_names))




#----build similarity links by 1)comparing each publisher name to every other name,2)score similarity, 3) if >= 85 the names become linked and stored-----------

threshold=85
links=[]

for i,name1 in enumerate(unique_publisher_names):
    for name2 in unique_publisher_names[i+1:]:
        score=fuzz.ratio(name1, name2)
        if score>=threshold:
            links.append((name1, name2))



#----build custers based off these links, to create a 'family'---------------------------------------------------------------------------------


parent= {name:name for name in unique_publisher_names}


#-----looks at each publisher name and finds what group it belongs to.
# ---returns the top level representative of that cluster with all subsequent names.

def find(x):
    while parent[x]!=x:
        parent[x]=parent[parent[x]]
        x=parent[x]
    return x


#--now that clusters have been created, some clusters may also be similar to other clusters, so they need to be merged.
#-----merges clusters together by comparing each leader for each name, comparing these two and merging based off whether they are linked.
def union(a,b):
    root_a=find(a)
    root_b=find(b)
    if root_a != root_b:
        parent[root_b]=root_a


#----repeats this function and applies it to every link. if A matches B, and B matches C,then all three can end up in the same cluster.
for a,b in links:
    union(a,b)


#----build the final grouped reuslts and store with the key being - cluster leader, and then all other publisher names that are in that cluster.
clusters={}
for name in unique_publisher_names:
    root=find(name)
    clusters.setdefault(root,[]).append(name)



#----create a new dataframe table for reference that contains all clusters------------------------
#-----this shows:
#cluster id
# cluster leader
# publisher name that is also in the category,
# the score to leader
# all variants of this publisher name 
# cluster size



rows = []
for cluster_id, (root, names) in enumerate(clusters.items(), start=1):
    cluster_leader = min(names, key=len)
    cluster_names = sorted(names)

    for name in cluster_names:
        score_to_leader = fuzz.ratio(name, cluster_leader)

        rows.append({
            "cluster_id": cluster_id,
            "cluster_threshold": threshold,
            "cluster_leader": cluster_leader,
            "name": name,
            "score_to_leader": score_to_leader,
            "all_variants": ", ".join(cluster_names),
            "cluster_size": len(cluster_names)})

#-----export it into a new file to reviw----------------------------------------

clusters_df = pd.DataFrame(rows)
clusters_df.to_csv(r"Outputs\publisher_name_clusters.csv", index=False)
