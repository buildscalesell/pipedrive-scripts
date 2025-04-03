install:
	python3 -m venv venv
	. venv/bin/activate && pip install --upgrade pip && pip install -r requirements.txt

run-clean:
	. venv/bin/activate && python clean_org_names.py

run-merge:
	. venv/bin/activate && python merge_similar_organizations.py

run-capitalize:
	. venv/bin/activate && python capitalize_org_names.py

