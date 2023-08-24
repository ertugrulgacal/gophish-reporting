from gophish import Gophish
import getopt, sys
import xlsxwriter

API_KEY = ""
api = Gophish(API_KEY, verify=False)

argumentList = sys.argv[1:]
options = "hli:o:"
long_options = ["help", "list", "id=", "output="]

try:
    arguments, values = getopt.getopt(argumentList, options, long_options)
    
    for currentArgument, currentValue in arguments:
        if currentArgument in ("-h", "--help"):
            print("arguments:")
            print("    -h, --help    show help message")
            print("    -l, --list    list all campaigns")
            print("    -i, --id      details about campaign with the given id")
            print("    -o, --output  write campaign details to file (only works when used with --id option)")
        
        elif currentArgument in ("-l", "--list"):
            for campaign in api.campaigns.get():
                print(f'Campaign ID:{campaign.id}\t\tCampaign Ismi:{campaign.name}\t\tCampaign Durumu:{campaign.status}')
                #print(campaign.results)
        
        elif currentArgument in ("-i", "--id"):
            campaignID = currentValue
            campaign = api.campaigns.get(campaignID)
            
            for result in campaign.results:
                print(f'id:{result.id}\t\tIsim:{result.first_name}\t\tSoyisim:{result.last_name}\t\tE-posta:{result.email}\t\tIP:{result.ip}\t\tDurum:{result.status}')
                
        elif currentArgument in ("-o", "--output"):
            fileName = currentValue + ".xlsx"
            
            workbook = xlsxwriter.Workbook(fileName)
            worksheet = workbook.add_worksheet()
            row = 1
            
            campaign = api.campaigns.get(campaignID)
            
            yellowBackground = workbook.add_format({'bg_color': 'yellow'})
            redBackground = workbook.add_format({'bg_color': 'red'})
            default1 = workbook.add_format({'bg_color': 'silver'})
            default2 = workbook.add_format({'bg_color': 'gray'})         
            
            worksheet.write(0, 0, "id", default2)
            worksheet.write(0, 1, "Isim", default2)
            worksheet.write(0, 2, "Soyisim", default2)
            worksheet.write(0, 3, "E-posta", default2)
            worksheet.write(0, 4, "IP", default2)
            worksheet.write(0, 5, "Durum", default2)
            
            for result in campaign.results:
                worksheet.write(row, 0, result.id, default1)
                worksheet.write(row, 1, result.first_name, default1)
                worksheet.write(row, 2, result.last_name, default1)
                worksheet.write(row, 3, result.email, default1)
                worksheet.write(row, 4, result.ip, default1)
                if result.status == "Clicked Link":
                    worksheet.write(row, 5, result.status, yellowBackground)
                elif result.status == "Submitted Data":
                    worksheet.write(row, 5, result.status, redBackground)
                else:
                    worksheet.write(row, 5, result.status, default1)
                              
                row += 1
                
            worksheet.autofit()
            workbook.close()
            
except getopt.error as err:
    # output error, and return with an error code
    print (str(err))
