/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import React, { Component } from 'react';
const axios = require('axios');


//const express = require('express')
//const app = express()
//var cors = require('cors');
//app.use(cors());


/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";


    document.getElementById("run").onclick = run;
  }
});

const resolveAfter2Seconds = async () => {
 /* var myHeaders = new Headers();
     myHeaders.append("Content-Type", "application/json");

     var raw = JSON.stringify({"email":"Dear All,\n Orel IT planned to conduct a React Js workshop for newly enlisted Software team.\n Date - 21st Nov 2020 Time - 1000 am to 0200 pm.Venue – Orel IT Nawinna, Maharagama.Therefore, I kindly request you all to physically participate to this event for get the maximum output. Here at Orel IT we are conducting this session adhere to the  maximum health protection and  I Kindly request you all to bring your own water bottles and hand sanitizers.Pls note – Bring your own laptops with React Js running environment. Any clarification pls contact - 0713249222.Thank you","options":"long"});

     var requestOptions = {
       method: 'POST', // *GET, POST, PUT, DELETE, etc.
       mode: 'no-cors', // no-cors, *cors, same-origin
       headers: {
         "Content-Type": "application/json",
         "Authorization": "Bearer sdfdsfs"
       },
       'Access-Control-Allow-Origin': '*',
      body: raw // body data type must match "Content-Type" header
     };

     const response = await fetch("https://spamcheck.postmarkapp.com/filter", requestOptions);

     const res = await response.json();

     console.log(res);*/
//   return new Promise(resolve => {
//     /*setTimeout(() => {
//       resolve('resolved');
      
//     }, 5000);*/




var myHeaders = new Headers();
myHeaders.append('Content-Type', 'application/json');
myHeaders.append('Access-Control-Allow-Origin', '*');
myHeaders.append("Access-Control-Allow-Credentials", "true");
myHeaders.append("Access-Control-Allow-Methods", "GET,HEAD,OPTIONS,POST,PUT");
myHeaders.append("Access-Control-Allow-Headers", "Access-Control-Allow-Headers, Origin,Accept, Authorization, X-Requested-With, Content-Type, Access-Control-Request-Method, Access-Control-Request-Headers");


var raw = JSON.stringify({"email":"Dear All,\n Orel IT planned to conduct a React Js workshop for newly enlisted Software team.\n Date - 21st Nov 2020 Time - 1000 am to 0200 pm.Venue – Orel IT Nawinna, Maharagama.Therefore, I kindly request you all to physically participate to this event for get the maximum output. Here at Orel IT we are conducting this session adhere to the  maximum health protection and  I Kindly request you all to bring your own water bottles and hand sanitizers.Pls note – Bring your own laptops with React Js running environment. Any clarification pls contact - 0713249222.Thank you","options":"long"});

var requestOptions = {
  method: 'POST',
  headers: myHeaders,
  body: raw,
  cors: 'no-cors',
  redirect: 'follow',

};

fetch("https://spamcheck.postmarkapp.com/filter", requestOptions)
  .then(response => resolve('response.text()'))
  .then(result => resolve('response.text()'))
  .catch(error => reject('error'));

//fog
//clone silla


//  /*     var myHeaders = new Headers();
//       myHeaders.append("Content-Type", "application/json");

//       var raw = JSON.stringify({"email_text":"Dear Tasdik, I would like to immediately transfer 10000 thousand dollars to your account as my beloved husband has expired and I have nobody to ask for to transfer the money to your account. I come from the family of the royal prince of burkino fasa and I would be more than obliged to take your help on this matter. Would you care to share your bank account details with me in the next email conversation that we have? -regards -Liah herman"});

//       var requestOptions = {
//         method: 'POST', // *GET, POST, PUT, DELETE, etc.
//         mode: 'no-cors', // no-cors, *cors, same-origin
//         cache: 'no-cache', // *default, no-cache, reload, force-cache, only-if-cached
//         credentials: 'same-origin', // include, *same-origin, omit
//         headers: {
//           'Content-Type': 'application/json'
//           // 'Content-Type': 'application/x-www-form-urlencoded',
//         },
//         redirect: 'follow', // manual, *follow, error
//         referrerPolicy: 'no-referrer', // no-referrer, *no-referrer-when-downgrade, origin, origin-when-cross-origin, same-origin, strict-origin, strict-origin-when-cross-origin, unsafe-url
//         body: raw // body data type must match "Content-Type" header
//       };

//      await fetch("https://plino.herokuapp.com/api/v1/classify/", requestOptions)*/






//      var myHeaders = new Headers();
//      myHeaders.append("Content-Type", "application/json");

//      var raw = JSON.stringify({"email":"Dear All,\n Orel IT planned to conduct a React Js workshop for newly enlisted Software team.\n Date - 21st Nov 2020 Time - 1000 am to 0200 pm.Venue – Orel IT Nawinna, Maharagama.Therefore, I kindly request you all to physically participate to this event for get the maximum output. Here at Orel IT we are conducting this session adhere to the  maximum health protection and  I Kindly request you all to bring your own water bottles and hand sanitizers.Pls note – Bring your own laptops with React Js running environment. Any clarification pls contact - 0713249222.Thank you","options":"long"});

//      var requestOptions = {
//        method: 'POST', // *GET, POST, PUT, DELETE, etc.
//        mode: 'no-cors', // no-cors, *cors, same-origin

//        headers: myHeaders,
//       body: raw // body data type must match "Content-Type" header
//      };

//     fetch("https://spamcheck.postmarkapp.com/filter", requestOptions)
//     .then(function (response) {
//       console.log(response.json());
//       //resolve(response.body)
//     })
//     .catch(function (error) {
//       console.log(error);
//       //resolve(error)
//     });


//     /*var raw = JSON.stringify({"email_text":"Dear Tasdik, I would like to immediately transfer 10000 thousand dollars to your account as my beloved husband has expired and I have nobody to ask for to transfer the money to your account. I come from the family of the royal prince of burkino fasa and I would be more than obliged to take your help on this matter. Would you care to share your bank account details with me in the next email conversation that we have? -regards -Liah herman"});


//     fetch('https://plino.herokuapp.com/api/v1/classify/', {
//       method: 'POST',
//       body: raw,
//       headers: {
//         'Content-type': 'application/json; charset=UTF-8'
//       }
//     })
//     .then(res => res.json())
//     .then(console.log)*/


// /*
//     var data = JSON.stringify({"email_text":"Dear Tasdik, I would like to immediately transfer 10000 thousand dollars to your account as my beloved husband has expired and I have nobody to ask for to transfer the money to your account. I come from the family of the royal prince of burkino fasa and I would be more than obliged to take your help on this matter. Would you care to share your bank account details with me in the next email conversation that we have? -regards -Liah herman"});

//     var config = {
//       method: 'post',
//       url: 'https://plino.herokuapp.com/api/v1/classify/',
//       mode: 'no-cors', // no-cors, *cors, same-origin
//       cache: 'no-cache', // *default, no-cache, reload, force-cache, only-if-cached
//       credentials: 'same-origin', // include, *same-origin, omit
//       headers: { 
//         'Content-Type': 'application/json',
//         'Access-Control-Allow-Origin': '*',
//         'Access-Control-Allow-Headers': '*',
//       },
//       data : data
//     };

//     axios(config)
//     .then(function (response) {
//       //console.log(JSON.stringify(response.data));
//       resolve(response)
//     })
//     .catch(function (error) {
//       console.log(error);
//       resolve(error)
//     });*/








//       //resolve('fdfdfdf')
//   });
}

async function asyncCall() {
  console.log('calling');
  //const result = 
  //console.log(result);
  return(await resolveAfter2Seconds())
  // expected output: "resolved"
}


export async function run() {
  /**
   * Insert your Outlook code here
   */


   
var mailItem = Office.context.mailbox.item;

mailItem.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
        var normalizeValue = null;

        if (result.value) {
            normalizeValue = result.value.split();

            
            
        }

        if (normalizeValue !== '') {
            mailItem.body.getAsync(Office.CoercionType.Html, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    // the value will be initialized in input value
                    



                    /*var data = JSON.stringify({"email_text":"Dear Tasdik, I would like to immediately transfer 10000 thousand dollars to your account as my beloved husband has expired and I have nobody to ask for to transfer the money to your account. I come from the family of the royal prince of burkino fasa and I would be more than obliged to take your help on this matter. Would you care to share your bank account details with me in the next email conversation that we have? -regards -Liah herman"});

                    var xhr = new XMLHttpRequest();
                    xhr.withCredentials = true;
                    
                    xhr.addEventListener("readystatechange", function() {
                      if(this.readyState === 4) {
                        console.log(this.responseText);
                      }
                    });
                    
                    xhr.open("POST", "https://plino.herokuapp.com/api/v1/classify/");
                    xhr.setRequestHeader("Content-Type", "application/json");
                    
                    let lll = xhr.send(data);
                    
                    document.getElementById("why").innerHTML = "<b>Subject:</b> <br/>" + result.value.trim();*/
                    
                    //asyncCall().then(re => document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + re)

                    

                   // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + result.value.trim();
                    //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + result.status();
                    //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + result.headers;
                    
                    //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + result.value.search("Subject");
                    

                    let vel = result.value.toLowerCase();
                    
                    let pos = vel.search("subject");
                    let header = vel.search("header");
                    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + pos;

                    // vel.forEach((ele) => {
                    //   if(ele.search("Subject") > 0){
                    //     document.getElementById("item-subject").innerHTML += ("<b>Subject:</b> <br/>" + ele);
                    //   }
                    // })


                    
                    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + result.value;

                    


                }
            });
        }
    }
});







}
