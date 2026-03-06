
.PHONY: unzip classify pack build delete

build: unzip delete classify pack

unzip:
	python unzip_in_folder.py ..

delete:
	cd .. && find . -type f ! \( -name '*.xls' -o -name '*.xlsx' -o -name '*.doc' -o -name '*.docx' -o -name '*.zip' -o -name '*.rar' \) -delete   
classify:
	python classify_md_by_name.py .. --move

pack:
	cd .. && find . -mindepth 1 -maxdepth 1 -type d ! -name 'CYLCollect' -exec zip -r "教育实践组.zip" {} +
