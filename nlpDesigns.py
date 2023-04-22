import pandas as pd
import spacy

# Load the spaCy English model
nlp = spacy.load('en_core_web_sm')

# Load the specs and designs data from separate tabs in the Excel sheet
specs_df = pd.read_excel('path/to/excel_file.xlsx', sheet_name='specs')
designs_df = pd.read_excel('path/to/excel_file.xlsx', sheet_name='designs')

# Preprocess the specs and designs text
specs_text = ' '.join(specs_df['specs Column Name'].tolist())
designs_text = ' '.join(designs_df['designs Column Name'].tolist())

# Analyze the preprocessed text with spaCy
specs_doc = nlp(specs_text)
designs_doc = nlp(designs_text)

# Extract relevant information from the specs and designs documents with spaCy
specs_nouns = set([token.lemma_ for token in specs_doc if token.pos_ == 'NOUN'])
specs_verbs = set([token.lemmait_ for token in specs_doc if token.pos_ == 'VERB'])
specs_adjectives = set([token.lemma_ for token in specs_doc if token.pos_ == 'ADJ'])
specs_entities = set([ent.text for ent in specs_doc.ents])

designs_nouns = set([token.lemma_ for token in designs_doc if token.pos_ == 'NOUN'])
designs_verbs = set([token.lemma_ for token in designs_doc if token.pos_ == 'VERB'])
designs_adjectives = set([token.lemma_ for token in designs_doc if token.pos_ == 'ADJ'])
designs_entities = set([ent.text for ent in designs_doc.ents])

# Compare the extracted information from the specs and designs documents
noun_similarity = len(specs_nouns.intersection(designs_nouns)) / len(specs_nouns.union(designs_nouns))
verb_similarity = len(specs_verbs.intersection(designs_verbs)) / len(specs_verbs.union(designs_verbs))
adj_similarity = len(specs_adjectives.intersection(designs_adjectives)) / len(specs_adjectives.union(designs_adjectives))
entity_similarity = len(specs_entities.intersection(designs_entities)) / len(specs_entities.union(designs_entities))

# Compute an overall similarity score
similarity_score = (noun_similarity + verb_similarity + adj_similarity + entity_similarity) / 4

# Write the similarity score to a new column in the designs sheet of the Excel file
designs_df['Similarity Score'] = similarity_score

# Save the updated Excel file
writer = pd.ExcelWriter('path/to/excel_file.xlsx')
specs_df.to_excel(writer, sheet_name='specs', index=False)
designs_df.to_excel(writer, sheet_name='designs', index=False)
writer.save()
