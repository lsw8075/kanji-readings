
import csv
import unicodedata

def sep_mandarin(mandarin):
    if mandarin == '':
        return ('', '')
    umap = {
        'u' : 'ü',
        'ū' : 'ǖ',
        'ú' : 'ǘ',
        'ǔ' : 'ǚ',
        'ù' : 'ǜ'
    }
    if mandarin[0] == 'w':
        if mandarin[1] in 'uūúǔù':
            return ('', mandarin[1:])
        else:
            return ('', 'u' + mandarin[1:])
    if mandarin[0] == 'y':
        if mandarin[1] in 'iīíǐì':
            return ('', mandarin[1:])
        elif mandarin[1] in 'uūúǔù':
            return ('', umap[mandarin[1]] + mandarin[2:])
        else:
            return ('', 'i' + mandarin[1:])
    if mandarin[0] in 'jqx' and mandarin[1] in 'uūúǔù':
        return (mandarin[0], umap[mandarin[1]] + mandarin[2:])
    index = 0
    for ch in mandarin:
        if ch in 'aāáǎàeēéěèiīíǐìoōóǒòuūúǔùüǖǘǚǜ':
            break
        index = index + 1
    
    return (mandarin[:index], mandarin[index:])

def sep_cantonese(cantonese):
    if cantonese == '':
        return ('', '')
    if cantonese == 'm' or cantonese == 'ng':
        return ('', cantonese)
    if cantonese[0] == 'j' and (cantonese[1] == 'i' or cantonese[1] == 'y'):
        return ('', cantonese[1:])
    if cantonese[0] == 'w' and cantonese[1] == 'u':
        return ('', cantonese[1:])
    index = 0
    for ch in cantonese:
        if ch in 'aeiouy':
            break
        index = index + 1
    
    return (cantonese[:index], cantonese[index:])

def sep_japanese(japanese):
    if japanese == '':
        return ('', '')
    if len(japanese) >= 2:
        if japanese[1] == 'h':
            if japanese[0] == 's':
                if japanese[2] == 'i':
                    return ('s', 'i')
                else:
                    return ('s', 'y' + japanese[2])
            elif japanese[0] == 'c':
                if japanese[2] == 'i':
                    return ('t', 'i')
                else:
                    return ('t', 'y' + japanese[2])
        elif japanese[0] == 't' or japanese[1] == 's':
            return ('t', 'u')
    if japanese[0] == 'j':
        if japanese[1] == 'i':
            return ('z', 'i')
        else:
            return ('z', 'y' + japanese[1])
    elif japanese[0] == 'f':
        return ('h', 'u')
    elif japanese[0] in 'kstnhmrgzdbp':
        return (japanese[:1], japanese[1:])
    return ('', japanese)

def sep_vietnamese(vietnamese):
    if vietnamese == '':
        return ('', '')
    if len(vietnamese) >= 2:
        if len(vietnamese) >= 3 and vietnamese[2] == 'h':
            return ('ng', vietnamese[3:])
        if vietnamese[1] == 'h' or vietnamese[1] == 'r':
            if vietnamese[0] == 'g':
                return ('g', vietnamese[2:])
            else:
                return (vietnamese[:2], vietnamese[2:])
    if vietnamese[0] == 'q' or vietnamese[0] == 'k':
        return ('c', vietnamese[1:])
    elif vietnamese[0] in 'bcdđghlmnprstvx':
        return (vietnamese[:1], vietnamese[1:])
    return ('', vietnamese)

def sep_korean(korean):
    if korean == '':
        return ('', '')
    nfd = unicodedata.normalize('NFD', korean)
    return (nfd[0], chr(0x110B) + nfd[1:])
    
unidata = dict()

def myzip(dual):
    first = ''
    second = ''
    for (inital, final) in dual:
        if inital == '' and final != '':
            first = first + ' ' + '/'
        else:
            first = first + ' ' + inital
        second = second + ' ' + final
    return (first[1:], second[1:])

with open('Unihan_Readings.txt') as unireadings:
    for line in unireadings.readlines():
        row = line.rstrip().split('\t')
        dataindex = chr(int(row[0][2:], 16))
        if dataindex not in unidata.keys():
            unidata[dataindex] = dict()
        unidata[dataindex][row[1]] = row[2]

with open('Unihan_DictionaryLikeData.txt') as unidic:
    for line in unidic.readlines():
        row = line.rstrip().split('\t')
        dataindex = chr(int(row[0][2:], 16))
        if dataindex not in unidata.keys():
            unidata[dataindex] = dict()
        unidata[dataindex][row[1]] = row[2]

with open('kw.csv', newline='') as csvfile:
    kwreader = csv.reader(csvfile, delimiter=' ')
    for row in kwreader:
        row = row.__str__().split(',')
        kanji = row[6]
        if kanji not in unidata.keys():
            unidata[kanji] = dict()
        final = row[3][1:-1][::-1]
        try:
            unidata[kanji]['inital'] = row[3][0]
            unidata[kanji]['final'] = final
            unidata[kanji]['tone'] = row[3][-1]
            if row[5] != '':
                unidata[kanji]['upper'] = row[5][0]
                unidata[kanji]['lower'] = row[5][1]
            elif row[4] != '':
                unidata[kanji]['upper'] = row[4][0]
                unidata[kanji]['lower'] = row[4][1]
            else:
                unidata[kanji]['upper'] = ''
                unidata[kanji]['lower'] = ''
        except:
            print(row)

mid = []

for k, v in unidata.items():
    output = [k, v.get('inital', ''), v.get('final', ''), v.get('tone', ''),
    v.get('upper', ''), v.get('lower', ''), 
    v.get('kHanyuPinyin', '').partition(':')[2].replace(',', ' '), 
    v.get('kCantonese', ''), v.get('kJapaneseOn', '').lower(), v.get('kVietnamese', ''),
    v.get('kHangul', '').translate(str.maketrans({'0':'', '1':'', '2':'', '3':'', 'E':'', 'N':'', 'X':'', ':':''})),
    v.get('kPhonetic', '').replace('*', '')]
    mid.append(output)

with open('result.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile, delimiter=',',
                            quotechar='|', quoting=csv.QUOTE_MINIMAL)
    writer.writerow(['kanji', 'inital', 'final', 'tone', 'upper', 'lower',
        'Minital', 'Mfinal', 'Mandarin', 'Cinital', 'Cfinal', 'Cantonese',
        'Jinital', 'Jfinal', 'Japanese', 'Vinital', 'Vfinal', 'Vietnamese',
        'Kinital', 'Kfinal', 'Korean', 'Pictophonetic'])
    for data in mid:
        man = myzip([sep_mandarin(x) for x in data[6].split(' ')])
        can = myzip([sep_mandarin(x) for x in data[7].split(' ')])
        jap = myzip([sep_japanese(x) for x in data[8].split(' ')])
        vet = myzip([sep_vietnamese(x) for x in data[9].split(' ')])
        kor = myzip([sep_korean(x) for x in data[10].split(' ')])
        dlist = [
            man[0], man[1], data[6],
            can[0], can[1], data[7],
            jap[0], jap[1], data[8],
            vet[0], vet[1], data[9],
            kor[0], kor[1], data[10]
        ]
        writer.writerow(data[0:6] + dlist + [data[11]])

