import time
import progressbar

data = []
count = 0
bar = progressbar.ProgressBar(max_vlaue=1000000) #Blue 'ProgressBar' is an object, not function
with open('reviews.txt','r') as f:
	for line in f:
		data.append(line)
		count += 1 # count = count + 1
		bar.update(count)
		# if count % 100000 == 0:
		# 	print(len(data)) # find the length(count of items) of list
print('Loading complete, total ', len(data), ' rows of record')

# List
sum_len = 0
for d in data:   # Set d as the length of each comment
	sum_len = sum_len + len(d)#

print('On average the length of comment is ', sum_len/len(data))

new = []
for d in data:
	if len(d) < 100:
		new.append(d)
print(len(new), ' comments in total less than 100 characters')
print(new[0])#


good = []
for d in data:
	if 'good' in d:
		good.append(d)
print(len(good), ' comments in total mentioned good')

#List comprehension
#good = [d for d in data if 'good' in d]
#print(good)#

#bad = ['bad' in d for d in data]
#print(bad)

# Original version
# bad = []
# for d in data:
# 	bad.append('bad' in d)


# Count word

start_time = time.time()
wc = {} # word_count
for d in data:
	words = d.split(' ')
	for word in words:
		if word in wc:
			wc[word] += 1
		else:
			wc[word] = 1 # add new key into dict

for word in wc:
	sorted_d = dict (sorted(wc.items(), key=lambda item: item[1],reverse=True))
	if wc[word] > 1000000:
		print(word, wc[word])
end_time = time.time()
print('It took', end_time - start_time, 'seconds')
print(len(wc))
print(wc['Allen'])

while True:
	word = input('Which word do you want to look up?')
	if word == 'q':
		break
	if word in wc:
		print(word, 'appears ', wc[word], ' times')
	else: 
		print('This word does not exist.')

print('Thanks for using this app!')












# print(data[0])
# print('----------------')
# print(data[1])

sum_len = 0
for d in data:   # Set d as the length of each comment
	sum_len = sum_len + len(d)#

print('On average the length of comment is ', sum_len/len(data))

new = []
for d in data:
	if len(d) < 100:
		new.append(d)
print(len(new), ' comments in total less than 100 characters')
print(new[0])#


good = []
for d in data:
	if 'good' in d:
		good.append(d)
print(len(good), ' comments in total mentioned good')


#List comprehension
#good = [d for d in data if 'good' in d]
#print(good)#

#bad = ['bad' in d for d in data]
#print(bad)

# Original version
# bad = []
# for d in data:
# 	bad.append('bad' in d)


