const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expected Realizations")
function dataExtraction_RE(graphql)
{
  var requestOptions = {
    'method': 'post',
    'payload': graphql,
    'contentType':'application/json',
    'headers':{
      'access_token': "7b27a8842839d827228622cb3608e88f884d8bc02e4b41cb907c3299c461793d"
    }
  };
  var response = UrlFetchApp.fetch(`https://gis-api.aiesec.org/graphql?access_token=${requestOptions["headers"]["access_token"]}`, requestOptions);
  var recievedDate = JSON.parse(response.getContentText());
  return recievedDate.data.allOpportunityApplication.data;
}

// Take the raw data recieved from the HTTP response and arrange it into the corresponding sheet
function dataUpdating()
{
      var queryAPDs = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n  date_approved:{from:"01/07/2022"}\n      \n\t\t}\n    \n    page:1\n    per_page:5000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n{\n created_at \tlc_alignment {\n\t\t\t\t\tkeywords\n\t\t\t\t}\n       id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n title \n\n\n  sub_product{\n\t\t\t\t\tname\n\t\t\t}\n\n     id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}`
      
    var graphql_APDs = JSON.stringify({query: queryAPDs})
    var dataSet_APDs = dataExtraction_RE(graphql_APDs);
    console.log("Data has been extracted")
    if(dataSet_APDs.length==0){
      return
    }

  var rows = [];
  var dataSet = dataSet_APDs

  var ids = sheet.getRange(1,1,sheet.getLastRow(),1).getValues().flat(1)

 for(var i = 0; i < dataSet.length; i++)
 {
      if(ids.indexOf(dataSet[i].person.id+"_"+dataSet[i].opportunity.id)>-1){
            var row = []
            row.push([
            dataSet[i].opportunity !=null?dataSet[i].person.id+"_"+dataSet[i].opportunity.id:dataSet[i].person.id+"_",
            dataSet[i].person.full_name,
            
            dataSet[i].status,
            dataSet[i].opportunity !=null?dataSet[i].opportunity.remote_opportunity == true?"Yes":"No":"",
            dataSet[i].person.email,  
            dataSet[i].person.contact_detail != null?dataSet[i].person.contact_detail.phone: "", 
            dataSet[i].person.home_mc.name, 
            dataSet[i].person.home_lc.name,
            dataSet[i].opportunity !=null? dataSet[i].opportunity.home_mc.name:"",
            dataSet[i].opportunity !=null?dataSet[i].opportunity.host_lc.name:"",
            dataSet[i].opportunity !=null?"https://expa.aiesec.org/opportunities/"+dataSet[i].opportunity.id:"",
            dataSet[i].opportunity !=null?dataSet[i].opportunity.programme.short_name_display:"",
            dataSet[i].opportunity !=null?dataSet[i].opportunity.opportunity_duration_type != null?dataSet[i].opportunity.opportunity_duration_type.duration_type:"":""/*duration_type*/,
            
            dataSet[i].date_approved != null?dataSet[i].date_approved.toString().substring(0,10):dataSet[i].updated_at.toString().substring(0,10)/*APD Date*/,
            dataSet[i].slot!=null?dataSet[i].slot.start_date:""/*slot start date*/,
            dataSet[i].slot!=null?dataSet[i].slot.end_date:""/*slot end date*/,
            dataSet[i].date_realized != null?dataSet[i].date_realized.toString().substring(0,10):"" /*RE date*/,
            dataSet[i].experience_end_date != null? dataSet[i].experience_end_date.toString().substring(0,10):""/*Fi date*/,
            dataSet[i].status == "remote_realized"?dataSet[i].updated_at.toString().substring(0,10):""/*remote date*/,
            dataSet[i].opportunity.sub_product != null? dataSet[i].opportunity.sub_product.name:"",
            dataSet[i].opportunity.title,
            dataSet[i].person.lc_alignment ? dataSet[i].person.lc_alignment.keywords:"-",
            dataSet[i].person.created_at
        ]);
        console.log(dataSet[i].opportunity.sub_product != null? dataSet[i].opportunity.sub_product.name:"")
        var index = ids.indexOf(dataSet[i].person.id+"_"+dataSet[i].opportunity.id)+1
        sheet.getRange(index,1,1,row[0].length).setValues(row)

      }
      else{
        console.log(i)
        console.log("old")
        rows.push([
            dataSet[i].opportunity !=null?dataSet[i].person.id+"_"+dataSet[i].opportunity.id:dataSet[i].person.id+"_",
            dataSet[i].person.full_name,
            dataSet[i].status,
            dataSet[i].opportunity !=null?dataSet[i].opportunity.remote_opportunity == true?"Yes":"No":"",
            dataSet[i].person.email,  
            dataSet[i].person.contact_detail != null?dataSet[i].person.contact_detail.phone: "", 
            dataSet[i].person.home_mc.name, 
            dataSet[i].person.home_lc.name,
            dataSet[i].opportunity !=null? dataSet[i].opportunity.home_mc.name:"",
            dataSet[i].opportunity !=null?dataSet[i].opportunity.host_lc.name:"",
            dataSet[i].opportunity !=null?"https://expa.aiesec.org/opportunities/"+dataSet[i].opportunity.id:"",
            dataSet[i].opportunity !=null?dataSet[i].opportunity.programme.short_name_display:"",
            dataSet[i].opportunity !=null?dataSet[i].opportunity.opportunity_duration_type != null?dataSet[i].opportunity.opportunity_duration_type.duration_type:"":""/*duration_type*/,
            dataSet[i].date_approved != null?dataSet[i].date_approved.toString().substring(0,10):dataSet[i].updated_at.toString().substring(0,10)/*APD Date*/,
            dataSet[i].slot!=null?dataSet[i].slot.start_date:""/*slot start date*/,
            dataSet[i].slot!=null?dataSet[i].slot.end_date:""/*slot end date*/,
            dataSet[i].date_realized != null?dataSet[i].date_realized.toString().substring(0,10):"" /*RE date*/,
            dataSet[i].experience_end_date != null? dataSet[i].experience_end_date.toString().substring(0,10):""/*Fi date*/,
            dataSet[i].status == "remote_realized"?dataSet[i].updated_at.toString().substring(0,10):""/*remote date*/,
            dataSet[i].opportunity.sub_product != null? dataSet[i].opportunity.sub_product.name:"",
            dataSet[i].opportunity.title,
            dataSet[i].person.lc_alignment ? dataSet[i].person.lc_alignment.keywords:"-",
            dataSet[i].person.created_at
        ]);
      }
 }
 if(rows.length > 0){
   sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows)
 }
  
}


function updateBreaks(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expected Realizations")
  var ids = sheet.getRange(1,1,sheet.getLastRow(),1).getValues()
  ids = ids.flat(1)
  var queryBreaks = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n       last_interaction:{from:\"01/07/2022\"}   \n\t\t\t       statuses:[\"approval_broken\",\"realization_broken\"]\n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    data\n    {\n      person\n      {\n        id\n      }\n      opportunity\n      {\n        id\n        }\n      \n      status      \n     }\n  }\n}`
  var graphql_Breaks = JSON.stringify({query: queryBreaks})
  var breaks = dataExtraction_RE(graphql_Breaks)
  for(var i = 243; i < breaks.length; i++){
    Logger.log(i)
    if(breaks[i].person != null && breaks[i].opportunity != null){
      var rowIndexInSheet = ids.indexOf(breaks[i].person.id+"_"+breaks[i].opportunity.id)  
      if(rowIndexInSheet != -1){
        sheet.getRange(rowIndexInSheet+1,3).setValue(breaks[i].status) 
      }
        
    }
  }
  

}



function getLastRow(sheet)
{ 
  var lr = sheet.getLastRow()
  var range = sheet.getRange(5,1,lr,1).getValues()
  var lastRow = 0;
  for(var i =0; i < lr; i++){
    if(range[i]!= "")
    {
      lastRow++;
    }
  }
  return lastRow;
}
