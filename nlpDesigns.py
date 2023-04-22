import pandas as pd
import spacy

# Load the spaCy English model
nlp = spacy.load('en_core_web_sm')

# Load the SRS and SDD data from separate tabs in the Excel sheet
srs_df = pd.read_excel('path/to/excel_file.xlsx', sheet_name='SRS')
sdd_df = pd.read_excel('path/to/excel_file.xlsx', sheet_name='SDD')

# Preprocess the SRS and SDD text
srs_text = ' '.join(srs_df['SRS Column Name'].tolist())
sdd_text = ' '.join(sdd_df['SDD Column Name'].tolist())

# Analyze the preprocessed text with spaCy
srs_doc = nlp(srs_text)
sdd_doc = nlp(sdd_text)

# Extract relevant information from the SRS and SDD documents with spaCy
srs_nouns = set([token.lemma_ for token in srs_doc if token.pos_ == 'NOUN'])
srs_verbs = set([token.lemma_ for token in srs_doc if token.pos_ == 'VERB'])
srs_adjectives = set([token.lemma_ for token in srs_doc if token.pos_ == 'ADJ'])
srs_entities = set([ent.text for ent in srs_doc.ents])

sdd_nouns = set([token.lemma_ for token in sdd_doc if token.pos_ == 'NOUN'])
sdd_verbs = set([token.lemma_ for token in sdd_doc if token.pos_ == 'VERB'])
sdd_adjectives = set([token.lemma_ for token in sdd_doc if token.pos_ == 'ADJ'])
sdd_entities = set([ent.text for ent in sdd_doc.ents])

# Compare the extracted information from the SRS and SDD documents
noun_similarity = len(srs_nouns.intersection(sdd_nouns)) / len(srs_nouns.union(sdd_nouns))
verb_similarity = len(srs_verbs.intersection(sdd_verbs)) / len(srs_verbs.union(sdd_verbs))
adj_similarity = len(srs_adjectives.intersection(sdd_adjectives)) / len(srs_adjectives.union(sdd_adjectives))
entity_similarity = len(srs_entities.intersection(sdd_entities)) / len(srs_entities.union(sdd_entities))

# Compute an overall similarity score
similarity_score = (noun_similarity + verb_similarity + adj_similarity + entity_similarity) / 4

# Write the similarity score to a new column in the SDD sheet of the Excel file
sdd_df['Similarity Score'] = similarity_score

# Save the updated Excel file
writer = pd.ExcelWriter('path/to/excel_file.xlsx')
srs_df.to_excel(writer, sheet_name='SRS', index=False)
sdd_df.to_excel(writer, sheet_name='SDD', index=False)
writer.save()
