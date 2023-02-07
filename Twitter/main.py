import requests
from datetime import datetime
from time import sleep
from openpyxl import Workbook

tweetsFound = 0
tweetData = []

#Apertura foglio excel
wb = Workbook()
ws = wb.active

def getResponse(queryString="", start_date=None,end_date=None,maxResultSinglerequest=10,maxTargetResult=10,next_token=None):
    global tweetsFound,tweetData,ws

    bearerToken = 'AAAAAAAAAAAAAAAAAAAAAPz%2FlQEAAAAAsGK8ys3ZwyN%2BZ9SNVn%2BLh3t%2F688%3D9HpIcNKogwwudG4yXpG0c8L882oUFLX3nIfKyz4uuQ93DQDgNq'
    endpoint = "https://api.twitter.com/2/tweets/search/all"

    params = {'query': queryString, "expansions": "author_id","tweet.fields": "created_at"}

    if start_date:
        start_date_format = datetime.strptime(start_date, '%d-%m-%Y').isoformat() + "Z"
        params['start_time'] = start_date_format
    if end_date:
        end_date_format = datetime.strptime(end_date, '%d-%m-%Y').isoformat() + "Z"
        params['end_time'] = end_date_format
    if maxResultSinglerequest != 10:
        if maxResultSinglerequest <= 500:
            params['max_results'] = maxResultSinglerequest
        else:
            params['max_results'] = 500
    if next_token:
        params['next_token'] = next_token
    
    response = requests.get(endpoint, params=params, headers={"Authorization": "Bearer "+bearerToken})


    if response.status_code == 200:
        response = response.json()
        tweetsFound += response['meta']['result_count']
        tweets = response["data"]
        new_next_token = ""
        if 'next_token' in response['meta']:
            new_next_token = response['meta']['next_token']
        print(next_token)
        for tweet in tweets:
            username = list(filter(lambda user: user['id'] == tweet['author_id'], response['includes']['users']))
            tweetData.append({'name': username[0]['name'], 'username': username[0]['username'], 'text': tweet['text'], 'date': datetime.strftime(datetime.fromisoformat(tweet['created_at']),"%d-%m-%Y")})
            ws.append((username[0]['name'], username[0]['username'], tweet['text'], datetime.strftime(datetime.fromisoformat(tweet['created_at']),"%d-%m-%Y")))
            wb.save('Tweets.xlsx')
        if tweetsFound < maxTargetResult:
            sleep(2)
            if new_next_token != "":
                getResponse(queryString, start_date=start_date,end_date=end_date,maxResultSinglerequest=maxResultSinglerequest,maxTargetResult=maxTargetResult,next_token=new_next_token)
    else:
        print(f"Request failed with status code: {response.status_code} \nError: {response.text}")

#Ricerca di tweet per Python che :
# contengano "Python"    
# siano di lingua inglese lang:en
# non siano risposte -is:reply          se si vogliono solo risposte is:reply senza il -
# maxResultSingleRequest è il massimo numero di tweet che una richiesta può darti. Il massimo che twitter mette a disposizione è 500
# maxTargetResult è il numero di tweet desiderati (NON E' DETTO CHE LO SCRIPT TERMINI COL NUMERO ESATTO DI TARGET RICHIESTO)
getResponse(queryString='Python "Python" lang:en -is:retweet -is:reply',start_date="12-01-2020", end_date="13-01-2020", maxResultSinglerequest=1000,maxTargetResult=10000)
