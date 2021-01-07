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
                    
                    var spam_words_arr=new Array(
                      "loan",
                      "winning",
                      "bulk email",
                      "mortgage",
                      "free",
                      "save",
                      "credit",
                      "amazing",
                      "bulk",
                      "opportunity",
                      "please read",
                      "reverses aging",
                      "hidden assets",
                      "stop snoring",
                      "free investment",
                      "dig up dirt on friends",
                      "stock disclaimer statement",
                      "multi level marketing",
                      "compare rates",
                      "cable converter",
                      "claims you can be removed from the list",
                      "removes wrinkles",
                      "compete for your business",
                      "free installation",
                      "free grant money",
                      "auto email removal",
                      "collect child support",
                      "free leads",
                      "amazing stuff",
                      "tells you it's an ad",
                      "cash bonus",
                      "promise you",
                      "claims to be in accordance with some spam law",
                      "search engine listings",
                      "free preview",
                      "act now! don't hesitate",
                      "credit bureaus",
                      "no investment",
                      "obligation",
                      "guarantee",
                      "refinance",
                      "price",
                      "affordable",
                      "home loan",
                      "lower your monthly payments",
                      "new low rate",
                      "Your Mortgage",
                      "Your refi",
                      "serious cash",
                      "personal",
                      "sale","cheap","deals","coupon","returned mail","refused to receive","returned to sender","cannot send message","attachment removed","infected message","undelivered mail","delivery status","virus infection","suspected spam","invite to join","banned content","mail scanner","Act now","Action","Apply now","Apply online","Buy","Buy direct","Call","Call now","Click here","Clearance","Click here","Do it today","Don’t delete","Drastically reduced","Exclusive deal","Expire","Get","Get it now","Get started now","Important information regarding","Instant","Limited time","New customers only","Now only","Offer expires","Once in a lifetime","Order now","Please read","Special promotion","Take action","This won’t last","Urgent","While stocks last","100%","All-new","Bargain","Best price","Bonus","Email marketing","Free","For instant access","Free gift","Free trial","Have you been turned down?","Great offer","Join millions of Americans","Incredible deal","Prize","Satisfaction guaranteed","Will not believe your eyes","As seen on","Click here","Click below","Deal","Direct email","Direct marketing","Do it today","Order now","Order today","Unlimited","What are you waiting for?","Visit our website","Acceptance","Access","Avoid bankruptcy","Boss","Cancel","Card accepted","Certified","Cheap","Compare","Compare rates","Congratulations","Credit card offers","Cures","Dear [personalization variable]","Dear friend","Drastically reduced","Easy terms","Free grant money","Free hosting","Free info","Free membership","Friend","Get out of debt","Giving away","Guarantee","Guaranteed","Have you been turned down?","Hello","Information you requested","Join millions","No age restrictions","No catch","No experience","No obligation","No purchase necessary","No questions asked","No strings attached","Offer","Opportunity","Save big","Winner","Winning","Won","You are a winner!","You’ve been selected!","Additional income","All-natural","Amazing","Be your own boss","Big bucks","Billion","Billion dollars","Cash","Cash bonus","Consolidate debt and credit","Consolidate your debt","Double your income","Earn","Earn cash","Earn extra cash","Eliminate bad credit","Eliminate debt","Extra","Fantastic deal","Financial freedom","Financially independent","Free investment","Free money","Get paid","Home","Home-based","Income","Increase sales","Increase traffic","Lose","Lose weight","Money back","No catch","No fees","No hidden costs","No strings attached","Potential earnings","Pure profit","Removes wrinkles","Reverses aging","Risk-free","Serious cash","Stop snoring","Vacation","Vacation offers","Weekend getaway","Weight loss","While you sleep","Work from home","Addresses","Beneficiary","Billing","Casino","Celebrity","Collect child support","Copy DVDs","Fast viagra delivery","Hidden","Human growth hormone","In accordance with laws","Investment","Junk","Legal","Life insurance","Loan","Lottery","Luxury car","Medicine","Meet singles","Message contains","Miracle","Money","Multi-level marketing","Nigerian","Offshore","Online degree","Online pharmacy","Passwords","Refinance","Request","Rolex","Score","Social security number","Spam","This isn’t spam","Undisclosed recipient","University diplomas","Unsecured credit","Unsolicited","US dollars","Valium","Viagra","Vicodin","Warranty","Xanax"
                      
                      ); 








                      
                      
                      let p = 'The quick brown one fox jumps over the lazy dog. If two the dog reacted, was it really lazy?';

                      const regexChar = /[^a-zA-Z0-9 ]/gi;
                      
                      const regmat = new RegExp('\\b(?:' + spam_words_arr.join('|') + ')\\b','gi');





                      //var re = RegExp("(?:^\\W*|(" + before.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")\\W+)" + error.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "(?:\\W+(" + after.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")|\\W*$)", "g"); 



                    //   function matchWords(subject, words) {
                    //     var regexMetachars = /[(){[*+?.\\^$|]/g;
                    
                    //     for (var i = 0; i < words.length; i++) {
                    //         words[i] = words[i].replace(regexMetachars, "\\$&");
                    //     }
                    
                    //     var regex = new RegExp("\\b(?:" + words.join("|") + ")\\b", "gi");
                    
                    //     return subject.match(regex) || [];
                    // }
                    
                    // matchWords(subject, ["one","two","three"]);

                      
                      try{
                        p = result.value.replace(regexChar, ' ');
                        //document.getElementById("item-subject").innerHTML = "<b>There are these spamy words included in the payload: </b> <br/>" + (result.value.match(regmat).length > 0)

                        if (result.value.match(regmat).length > 0){
                          document.getElementById("item-subject").innerHTML = "<b>There are these spamy words included in the payload: </b> <br/>" + result.value.match(regmat)
                          document.getElementById("item-subject2").innerHTML = "<b>This email might be an spam</b> <br/>"
                        }else{
                          document.getElementById("item-subject").innerHTML = "<b>There are no spamy words included in the payload:</b> <br/>" 
                        }
                      }catch(e){
                        //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + e;
                      }
                    //let vel = result.value.toLowerCase();
                    
                    // let pos = vel.search("subject");
                    // let header = vel.search("header");
                   // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + pos;

                    // vel.forEach((ele) => {
                    //   if(ele.search("Subject") > 0){
                    //     document.getElementById("item-subject").innerHTML += ("<b>Subject:</b> <br/>" + ele);
                    //   }
                    // })


                    
                    //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + result.value;

                    


                }
            });
        }
    }
});







}
