import re
string = 'abe(ac）ad）'
p1 = re.compile(r'[(](.*?)[)]', re.S) #最小匹配
p2 = re.compile(r'[(|（](.*)[)|）]', re.S)  #贪婪匹配
print(re.findall(p1, string))
print(re.findall(p2, string))