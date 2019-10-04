import pathlib
import pandas as pd
import datetime 
import gzip
import shutil
import glob
import os
import pyspark.sql.functions as Fun
import pyspark.sql.types as Type
import tarfile
from collections import OrderedDict
str = "log.!""#$%'()[]{}\|~^=-?_.@*+,<あVersion:B-First_Software_Log_1.00\tfunctionId:"
list =[OrderedDict(cell.split(':', 1) for cell in line.split('\t')) for line in str[:-1].split('\n')]
df = spark.createDataFrame(list)
df.show()




・相手が自分との壁を感じさせないコミュニケーションを習得するために、コミュニケーションをとる際は、
自分から”自分”の事を話す癖を業務内、外で身につける。


・データベーススキルに関して、後輩に教えれる立場になる、先輩方と同じ目線で話せる様になるために、
脱初級者向けの技術書を読破し、業務外で学んだスキルを業務に活かせるようになる。

・自分が面白くないなと思った仕事に対しても、面白みをみつけるために、柔軟な発想を持てるようになる。
柔軟な発想を持てるようになるために、様々な業界の本を読む、色々なことを体験することで、自分の可能性を広げる


・自分がタスクを行う際、頼まれた事を一方通行で行うのではなく、
複数の観点から自分が思う最善の方針を決める習慣を定着させる。
・方針が決まったら上司に相談し、アドバイスを柔軟に取り入れることで、自分が業務を行う上での手札を増やし、
後輩からの仕事の相談にアドバイスをできる準備をする


・お客様と密な関係を築くために、場面に応じた相手に寄り添える言葉遣い、話し方を成長させる。
そのために、定例MTG、お客様との打ち合わせ等を利用し、会議の目的を確認した上で先輩方の振る舞いから自分に足りないものをリストアップする。
→先輩方から学んだ技術を生かし、論点をずらさないコミュニケーションを心掛ける
