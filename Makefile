
.PHONY: unzip classify pack build

build: unzip classify pack

unzip:
	python unzip_in_folder.py ..

classify:
	python classify_md_by_name.py .. --move

pack:
	cd .. && find . -mindepth 1 -maxdepth 1 -type d ! -name 'CYLCollect' -exec zip -r "教育实践组.zip" {} +
