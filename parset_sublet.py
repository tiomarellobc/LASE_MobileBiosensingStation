def Parse_Channels(test):
    Top_Channels = test.split(",")
    All_Channels = []
    for Top_Channel in Top_Channels:
        if(":" in Top_Channel):
            Bwoah = Top_Channel.split(':')
            low = int(Bwoah[0])
            high = int(Bwoah[1])+1
            for Sub_Channels in range(low, high):
                All_Channels.append(str(Sub_Channels))
        else:
            All_Channels.append(Top_Channel)
    return(All_Channels)


channels = Parse_Channels("101:104,107, 110:115")
print(channels)
print(",".join(channels))