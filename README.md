# consul for 1C

### Установка на серверах приложений 1С 

`git clone -b master https://gitgub.com/farukshin/consul1c.git`

Скачать consul.exe с [официального сайта](https://www.consul.io/downloads.html) и положить в папку consul1c

В файле conf\conf.json в параметр start_join добавить IP адреса серверов приложений 1С, на которых планируется запуск consul

Создать службу consul, запускающую consul.exe с параметрами agent -config-dir="conf"
