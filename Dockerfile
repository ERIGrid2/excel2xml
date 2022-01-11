from alpine:latest

run apk update
run apk add python3 py3-pip
run /usr/bin/pip3 install pyxb openpyxl
copy . /excel2md
cmd sh /excel2md/process_all_xlsx.sh
