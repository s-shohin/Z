#!/usr/bin/env python
# coding: utf-8

# In[13]:


from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys

import pandas as pd
import openpyxl as xl
import re
import traceback


# In[14]:


def Z_func(data):
    USER_ID='nihonsaburo1234@gmail.com'
    USER_PW='s@fic1234'

    #dict型のdataを受け取り、打鍵結果を入力したdict型のdataを返す
    try:
        options = webdriver.ChromeOptions()
        #options.add_argument('--headless') #ブラウザ表示なし#Zはheadlessだとエラーになる。
        options.add_argument('--incognito') #シークレットモード 
        browser = webdriver.Chrome(options=options)

        #見積もりページを開く
        url= "https://www.zurich.co.jp/auto/common/ncdAssessmentPage.html"
        browser.get(url)
        sleep(1)

        if 'S' in data['NF2']:
            new = 'New'
            browser.find_element(By.CSS_SELECTOR, '#page1_question1_answer2 > div > span.txt > p.btnpara').click()#新規
            sleep(1)
            browser.find_element(By.CSS_SELECTOR, '#btnQ1_HasNoBusiness > a > div > dl > dt').click()#過去契約がない

            #セカンドカー
            if data['NF2']=='7S':
                browser.find_element(By.CSS_SELECTOR, '#btnQ2_HasAnotherCar > a > div > dl > dd').click()
                browser.find_element(By.CSS_SELECTOR, '#btnQ3_ApplySecondCar > a > div > dl > dt').click()
                browser.find_element(By.CSS_SELECTOR, '#part_newSecondCarBtnBlock > p > a').click()

            else:
                browser.find_element(By.CSS_SELECTOR, '#btnQ2_NoAnotherCar > a > div > dl > dt').click()
                browser.find_element(By.CSS_SELECTOR, '#part_newBusinessBtnBlock > p > a > span').click()
        else:
            new = ''
            browser.find_element(By.CSS_SELECTOR, '#page1_question1_answer1 > div > span.txt > p.btntxt > img').click()

        if 'S' in data['NF2']:
            pass
        else:
            #一年契約
            sleep(3)
            browser.find_element(By.CSS_SELECTOR, '#riskFactorForm\:selectYearType > tbody > tr > td.firstChild > label').click()   
        
        sleep(1)

        #始期日
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:commencementDateEraYearField')).select_by_visible_text(data['西暦2']) 
        sleep(2)
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:commencementDateMonthField')).select_by_visible_text(str(data['月2']))
        sleep(2)
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:commencementDateDayField')).select_by_visible_text(str(data['日2'])) 
        sleep(2)

        if 'S' in data['NF2']:
            pass
        else:
            #加入中の保険会社
            Select(browser.find_element(By.CSS_SELECTOR,'#riskFactorForm\:insuranceCompanyCodeField')).select_by_index(1)
            sleep(3)
            #現在の等級
            Select(browser.find_element(By.CSS_SELECTOR,'#riskFactorForm\:currentNoClaimDiscountField')).select_by_visible_text(data['NF2'])
            sleep(3)
            #事故件数
            browser.find_element(By.CSS_SELECTOR, '#riskFactorForm\:noPreviousClaimsField > tbody > tr > td.firstChild > label').click()
            sleep(1)

            #事故有係数
            browser.find_element(By.CSS_SELECTOR, '#riskFactorForm\:prevYrAccidentPeriodCountField > tbody > tr > td.firstChild > label').click()
            sleep(1)

            #車両保険あり
            browser.find_element(By.CSS_SELECTOR, '#riskFactorForm\:ownDamageCoverageField > tbody > tr > td.firstChild > label').click()
            sleep(2)                                         
            
            #運限
            if data['運限修正'] == '本配':
                browser.find_element(By.CSS_SELECTOR, '#riskFactorForm\:driverLimitationDiscountField > tbody > tr > td.nthChild2n.nthChild2 > label').click()
            elif data['運限修正'] == '家族':
                browser.find_element(By.CSS_SELECTOR, '#riskFactorForm\:driverLimitationDiscountField > tbody > tr > td.firstChild > label').click()
            else:
                browser.find_element(By.CSS_SELECTOR, '#riskFactorForm\:driverLimitationDiscountField > tbody > tr > td:nth-child(4) > label').click()
            sleep(1)

            #年齢
            Select(browser.find_element(By.CSS_SELECTOR,'#riskFactorForm\:ageLimitedDiscountField')).select_by_visible_text(data['年齢限定修正2'])
            #sleep(1)

        #初度登録
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:carRegistrationDateEraYearField')).select_by_visible_text(data['初度年2']) 
        sleep(1)
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:carRegistrationDateMonthField')).select_by_visible_text(str(data['初度月2']))
        sleep(2)

        #型式
        browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:carTypeField_input').send_keys(data['型式2'])
        sleep(2)
        browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:autoCompleteButton').click()
        sleep(2)
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:dummyCarTypeListField0')).select_by_index(1)
        sleep(1)

        #記名被保険者は契約者の配偶者男性
        browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:mainDriverRelationField > tbody > tr > td.nthChild3.nthChild3n > label').click()

        #生年月日
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:insuredDOBEraYearField')).select_by_visible_text(data['生年2']) 
        sleep(2)
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:insuredDOBMonthField')).select_by_visible_text(str(data['生まれ月2']))
        sleep(2)
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:insuredDOBDayField')).select_by_visible_text(str(data['生まれ日2'])) 
        sleep(2)
        Select(browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:mainDriverResidencialAreaField')).select_by_visible_text(str(data['地域2'])) 
        sleep(2)

        #免許の色
        if data['免許2'] == 'ゴールド':
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:driversLicenseColorAutoField > tbody > tr > td.firstChild > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:driversLicenseColorAutoField > tbody > tr > td.nthChild2 > label').click()

        sleep(1)
        #目的
        if data['使用目的2'] == '日常':
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:vehicleUsageField > tbody > tr > td.firstChild > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:vehicleUsageField > tbody > tr > td.nthChild3.nthChild3n > label').click()#通勤25km
        sleep(1)

        #目的
        if data['走行距離2'] == 1:
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:anualMilleageField > tbody > tr > td.firstChild > label').click()
        elif data['走行距離2'] == 2:
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:anualMilleageField > tbody > tr > td.nthChild2 > label').click()
        elif data['走行距離2'] == 3:
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:anualMilleageField > tbody > tr > td.nthChild3 > label').click()
        elif data['走行距離2'] == 4:
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:anualMilleageField > tbody > tr > td:nth-child(4) > label').click()
        elif data['走行距離2'] == 5:
            browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:anualMilleageField > tbody > tr > td:nth-child(5) > label').click()
        sleep(1)

        if 'S' in data['NF2']:
            #運限
            if data['運限修正'] == '本配':
                browser.find_element(By.CSS_SELECTOR, '#riskFactorNewForm\:driverLimitationDiscountField > tbody > tr > td.firstChild > label').click()
            elif data['運限修正'] == '家族':
                browser.find_element(By.CSS_SELECTOR, '#riskFactorNewForm\:driverLimitationDiscountField > tbody > tr > td.nthChild2n.nthChild2 > label').click()
            else:
                browser.find_element(By.CSS_SELECTOR, '#riskFactorNewForm\:driverLimitationDiscountField > tbody > tr > td:nth-child(3) > label').click()
            sleep(1)

            #年齢
            if data['年齢限定修正2'] == '全年齢補償':
                browser.find_element(By.CSS_SELECTOR, '#riskFactorNewForm\:ageLimitedDiscountField > tbody > tr > td.firstChild > label').click()
            elif data['年齢限定修正2'] == '21歳以上補償':
                browser.find_element(By.CSS_SELECTOR, '#riskFactorNewForm\:ageLimitedDiscountField > tbody > tr > td.nthChild2n.nthChild2 > label').click()
            elif data['年齢限定修正2'] == '26歳以上補償':
                browser.find_element(By.CSS_SELECTOR, '#riskFactorNewForm\:ageLimitedDiscountField > tbody > tr > td:nth-child(3) > label').click()
            else:
                browser.find_element(By.CSS_SELECTOR, '#riskFactorNewForm\:ageLimitedDiscountField > tbody > tr > td:nth-child(4) > label').click()
            sleep(1)

            #セカンドカー有りの場合、他の保険会社
            if len(browser.find_elements(By.CSS_SELECTOR,'#riskFactorNewForm\:insuranceCompanyCodeField'))>0:
                Select(browser.find_element(By.CSS_SELECTOR,'#riskFactorNewForm\:insuranceCompanyCodeField')).select_by_index(1)
        else:
            pass

        #何で知ったか
        browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:mediaField > tbody > tr > td.nthChild3.nthChild3n > label').click()
        sleep(1)

        #次へ
        browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:button > img').click()
        sleep(1)

        #はい
        browser.find_element(By.CSS_SELECTOR, '#riskFactor'+ new + 'Form\:yes_button').click()


        #ログイン画面に移動
        browser.find_element(By.CSS_SELECTOR, '#loginLinkBlock > p > a').click()
        #IDでログイン
        browser.find_element(By.CSS_SELECTOR, '#login-id_panel > span').click()
        #IDとPWを入力
        browser.find_element(By.CSS_SELECTOR, '#login_id').send_keys(USER_ID)
        browser.find_element(By.CSS_SELECTOR, '#password').send_keys(USER_PW)

        #ログイン
        browser.find_element(By.CSS_SELECTOR, '#login-id > div.str-container > div > form > p.mod-form-btn-bext').click()

        browser.execute_script("window.scrollTo(0, 10000)")#下までスクロールしないとエラーが起きがち？
        sleep(2)

        #次へ なぜかXPATHで指定しないと上手く押せない
        browser.find_element(By.XPATH, '/html/body/main/div[2]/form/ul/li[1]').click()


        sleep(1)
        #次へすすむ
        browser.find_element(By.CSS_SELECTOR, 'body > div.str_main > div.str_main_inner > form > ul > li._button').click()
        ##結果画面

        ##標準プラン

        #対物LL
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:propertyDifferenceExceededClause_plan1')).select_by_visible_text(data['対物LL2']) 

        browser.execute_script("window.scrollTo(0, 40000)")#適当にスクロール

        #人身傷害
        if data['人傷種類']=='車内外':
            browser.find_element(By.CSS_SELECTOR, '#limitedVolOnInsuredCarClausePanel_plan1 > ul > li.radio05 > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#limitedVolOnInsuredCarClausePanel_plan1 > ul > li.radio04 > label').click()

        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:personalInjuryOption_plan1')).select_by_visible_text(data['人傷AMT2']) 

        #搭傷
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:passengerPASumInsured_plan1')).select_by_visible_text(data['搭乗者2']) 

        #車両保険
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:ownDamageOption_plan1')).select_by_visible_text(data['車両保険種類2']) 
        try:
            Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:ownDamageSumInsured_plan1')).select_by_visible_text(data['車両AMT2']) 
        except:
            data['車両AMTエラー']='該当なし'
        sleep(3)
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:ownDamageSpecialClauseOption_plan1')).select_by_visible_text(data['車両免責2']) 

        #特約
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:personalEffectiveOption_plan1')).select_by_visible_text(data['積載動産2']) 

        #代車
        if data['代車費用2']=='あり':
            browser.find_element(By.CSS_SELECTOR, '#availableComplementaryCarPanel_plan1 > ul > li.radio01 > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#availableComplementaryCarPanel_plan1 > ul > li.radio02 > label').click()

        #地噴津
        if data['地噴津']=='あり':
            browser.find_element(By.CSS_SELECTOR, '#earthQuakeVehiclesClausePanel_plan1 > ul > li.radio01 > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#earthQuakeVehiclesClausePanel_plan1 > ul > li.radio02 > label').click()

        #その他の特約
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:personalLiabilityOption_plan1')).select_by_visible_text(data['個賠2']) 
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:familyBikeOption_plan1')).select_by_visible_text(data['原付2']) 
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:paOption_plan1')).select_by_visible_text(data['傷害2']) 
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:earthQuakeInsuredDeathClause_plan1')).select_by_visible_text(data['地噴津死亡傷害2']) 

        browser.execute_script("window.scrollTo(0, 60000)")#適当にスクロール

        if data['弁特2']=='あり':
            browser.find_element(By.CSS_SELECTOR, '#lawyerExpenseSecurityOptionPanel_plan1 > ul > li.radio01 > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#lawyerExpenseSecurityOptionPanel_plan1 > ul > li.radio02 > label').click()


        browser.execute_script("window.scrollTo(0, 0)")#上までスクロール

        sleep(3)

        #保険料再計算
        try:
            browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:recalculateBtn1_top > span > img').click()
        except:
            pass

        sleep(2)
        #車有保険料の取得
        result0_discount = browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:quotePremium1_top').text
        result0_discount_amt = browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:discountArea1_top').text

        ##プラン比較①
        browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:planComment2').click()

        #対物LL
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:propertyDifferenceExceededClause_plan2')).select_by_visible_text(data['対物LL2']) 

        browser.execute_script("window.scrollTo(0, 20000)")#適当にスクロール

        #人身傷害
        if data['人傷種類']=='車内外':
            browser.find_element(By.CSS_SELECTOR, '#limitedVolOnInsuredCarClausePanel_plan2 > ul > li.radio05 > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#limitedVolOnInsuredCarClausePanel_plan2 > ul > li.radio04 > label').click()

        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:personalInjuryOption_plan2')).select_by_visible_text(data['人傷AMT2']) 

        #搭傷
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:passengerPASumInsured_plan2')).select_by_visible_text(data['搭乗者2']) 

        #車両保険
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:ownDamageOption_plan2')).select_by_visible_text('なし') 

        #特約
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:personalEffectiveOption_plan2')).select_by_visible_text(data['積載動産2']) 

        #その他の特約
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:personalLiabilityOption_plan2')).select_by_visible_text(data['個賠2']) 
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:familyBikeOption_plan2')).select_by_visible_text(data['原付2']) 
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:paOption_plan2')).select_by_visible_text(data['傷害2']) 
        Select(browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:earthQuakeInsuredDeathClause_plan2')).select_by_visible_text(data['地噴津死亡傷害2']) 

        browser.execute_script("window.scrollTo(0, 40000)")#適当にスクロール

        if data['弁特2']=='あり':
            browser.find_element(By.CSS_SELECTOR, '#lawyerExpenseSecurityOptionPanel_plan2 > ul > li.radio01 > label').click()
        else:
            browser.find_element(By.CSS_SELECTOR, '#lawyerExpenseSecurityOptionPanel_plan2 > ul > li.radio02 > label').click()
        sleep(3)

        #保険料再計算
        try:
            browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:recalculateBtn2_top > span > img').click()
        except:
            pass

        sleep(2)

        #車無保険料の取得
        result1_discount = browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:quotePremium2_top').text
        result1_discount_amt = browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:discountArea2_top').text
        early_discount_amt = browser.find_element(By.CSS_SELECTOR, '#calculatePremiumForm\:discountArea2_early').text

        #空白なら、文字列の0をいれる。後段のre.subでエラーを起こさないように。
        if early_discount_amt == "":
            early_discount_amt = str(0)
       
        #カンマを除く
        data['車有P']=int(re.sub(r"\D", "", result0_discount))
        data['車無P']=int(re.sub(r"\D", "", result1_discount))
        data['イ割車有']=int(re.sub(r"\D", "", result0_discount_amt))
        data['イ割車無']=int(re.sub(r"\D", "", result1_discount_amt))
        data['早割']=int(re.sub(r"\D", "", early_discount_amt))
        #browser.quit()

    #不測のエラーが起きた場合は、結果にEを入力する
    except :
        data['車有P']='E'
        data['車無P']=traceback.format_exc()
        #browser.quit()

    return data  


# In[15]:


if __name__ == "__main__":
    FAIL_NAME='Z条件_データ1_定点'
    SHEET_NAME='Z打鍵'

    #データ読み込み
    df=pd.read_excel(FAIL_NAME+'.xlsm',sheet_name=SHEET_NAME)
    data=df.loc[14-2,:].to_dict()
    Z_func(data) 
    print(data)

