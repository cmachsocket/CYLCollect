
.PHONY: unzip classify

build: unzip classify

unzip:
	python unzip_in_folder.py ..

classify:
	python classify_md_by_name.py .. --move


