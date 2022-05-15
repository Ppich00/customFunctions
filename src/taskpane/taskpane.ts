/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
import * as platform from "@office-platform/office-web-service-client";
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
});

export async function run() {
  try {
    // await Excel.run(async (context) => {
      // const cell = context.workbook.getActiveCell();
      await data();
      // console.log(context);
    // });
  } catch (error) {
    console.error(error);
  }
}

export async function data() {
  platform.configure({
    platform: "beta",
  });
  const session = platform.createSession({ autoOpen: false });
  debugger;

  console.log("custom data");
  const token =
    "eyJ0eXAiOiJhdCtqd3QiLCJhbGciOiJSUzI1NiIsImtpZCI6InRhbTg2amltMDY4dTh0LWpPU1c3ZWhyTUx6VW1INVVhOS1JOGRtTzA3Zm8ifQ.eyJkYXRhIjoie1wiY2lwaGVydGV4dFwiOlwiNzJyMG15TV80TEJrM1luMDVPTFJ2OE1OT2xvUkRFempkQ3NROEFyY2wwWG9oNWM2UFo0UjVvTDNlVDBwTzdvVVhnQXRNcmtadjhET29KU2NyVTZqVFpBU09GcU03YmVUQUNXTXozTWpKQWdqc2tRSDUtazdPY21aOFVZMHdGekEtcWRNZkltd3U1ckxEWjBfemsyR2RBVzB1NlVLdUREeXpoTlo3dUUxbmhRWVUyR2tuNXlPSDJmYUFkeXhYamUyNmhxa1ZuRTVTZTVDTloyN2tHbDJBcm1DTDZPNUxhUVdzRzhMb0JRbF9kTWdSa2hjamYtYkJaTEtKYXB6MDRRdkY3SEcwWVhHU0RBZk9TdVdUbm5tZ3lKdHEtd1hlT3U5Y3hiaWR3cWdFLW1ranRFNXlxcFpqVEt4RkNxdWpIbHpYYkJfaWlTcUNsMTA5a25nYXBzU3o1aGozajVEZ3Yxc1RhQ0RRZThyczYyazJlcE5oMFc5RDFNcGtZZ1BEd25SVTNsN293cmR4Z1FLNm9YOEdSV1RQOVZRbUNONHdBRXFVclBjTTg2WVNwdWNtXzhZMXp2ZWZIOWQ4UlBMSWttYVFFcmZRZjdqM1hpblN4Mk16bkd4U2tzM1ZVZHM3NU9HbnhoZEY4X2xVOHZrU3NqNE9jVzNVbElmNFBoRzZwaWJ5RjRmLTRGNFJhX0s3R3FNYXFLT2tNWnp1dm9kc0J6dWZYRFI5cW44a3RrR2tSTkNrSzR3RThROElHY05rNDJTazNpV3lkbHNuNGtlSXVId3d3RUZfVU1kZ2ZXOFZfZmthTTBEMVpLQkFIcnk2VExGOXYwcUZPY2RwWmdZZ0lTbURQWTRvd3hUZGZCd2hKb1B6YkFOYWRaQkREb29wZEotMUZQNkVxZUlORkxZdXZpaU53dTZhN3B5T3hTc0ZMUEtoYmRBMDFCcGpkWGNQeXdOSzZoTVRPaWJQOHRjRl9FZkNma1lycjhYYVVVdUJKYm5ibjlvMEFnUVJKNTExVE9rc3piTmVnS0hlNnp4UjJLdU5SS0lrSmVCTno0OG1NVmRZVVQ2WWVNOVZLSGRnckplRGZHRE1GWjNMNkc0dUtsSFwiLFwiaXZcIjpcIldySkJ5T2FsdmlyQkVpaGlcIixcInByb3RlY3RlZFwiOlwiZXlKaGJHY2lPaUpCVjFOZlJVNURYMU5FUzE5Qk1qVTJJaXdpWlc1aklqb2lRVEkxTmtkRFRTSXNJbnBwY0NJNklrUkZSaUo5XCIsXCJyZWNpcGllbnRzXCI6W3tcImVuY3J5cHRlZF9rZXlcIjpcIkFRSUJBSGo5b2xkMjlFcHdVTWFlNkJKNFFVZWs4ZkxBOXFDa3diUEpQUXVBX20wR3BRRzM0UzRkbHJsYjc3U2tXbXAzaS1nTkFBQUFmakI4QmdrcWhraUc5dzBCQndhZ2J6QnRBZ0VBTUdnR0NTcUdTSWIzRFFFSEFUQWVCZ2xnaGtnQlpRTUVBUzR3RVFRTUs3MHlZZzZrY3dKX244ejhBZ0VRZ0RzXzB6NlZSWmJtN0RWT240YWJjTlZPLXhNWXNYX1Z4TUhQNnRIWkx4Wl95RlBRZGpkbmp4OXk2V0NvdGI4czlqeEVNelY0dFZkUXdaOUJld1wiLFwiaGVhZGVyXCI6e1wia2lkXCI6XCJhcm46YXdzOmttczp1cy1lYXN0LTE6NjUzNTUxOTcwMjEwOmtleS82NTA4ZDMzZi0yNTgzLTRiM2YtOWRmMi1iZDBlNDRjNTFlNDNcIn19LHtcImVuY3J5cHRlZF9rZXlcIjpcIkFRSUNBSGpsVjNnZWxrcTBsYm1IRHJKbFNwUExpS2U0ajNjWngyS210d3YwVHZxWUxnRXVqbFNTYzdJMFZuXzJfdFdjdHpfS0FBQUFmakI4QmdrcWhraUc5dzBCQndhZ2J6QnRBZ0VBTUdnR0NTcUdTSWIzRFFFSEFUQWVCZ2xnaGtnQlpRTUVBUzR3RVFRTVQ1LTE3RmM1MzZ1X0NEQmZBZ0VRZ0R1aHFFVEhuemI5T2haVkFST2cwN0lSem03SGMxOTdic2U4dUV3RHdDM052YTBlMjN5SDR2d29Pd1h3UGNqbk94eVRuSWxneHFjODJTdHBwZ1wiLFwiaGVhZGVyXCI6e1wia2lkXCI6XCJhcm46YXdzOmttczphcC1zb3V0aGVhc3QtMTo2NTM1NTE5NzAyMTA6a2V5LzYyYTM2ODM1LWJjYjUtNDdmOS1iYjA1LTc5OTIwNTdjMTVjN1wifX0se1wiZW5jcnlwdGVkX2tleVwiOlwiQVFJQ0FIaFhLUTNoOWFPTVU4T19uLWctSElfd0pmME5kQkc3NER6S0owM1Q5NE9nd1FFSk9yS21uamNablpBVDJkWXUtT0IwQUFBQWZqQjhCZ2txaGtpRzl3MEJCd2FnYnpCdEFnRUFNR2dHQ1NxR1NJYjNEUUVIQVRBZUJnbGdoa2dCWlFNRUFTNHdFUVFNSTdaQ242WVcwb0lDMEZlR0FnRVFnRHZnTzFrTkVIcmNmZFEycHRQa1pxbzhEbmVJMDU2MHp5UzFqaklxOWh0UUprTWJMT1daT1BMTUVSOHl6eGVQd2tHMm9HNFFJLTdKRUVmSWJ3XCIsXCJoZWFkZXJcIjp7XCJraWRcIjpcImFybjphd3M6a21zOmV1LXdlc3QtMTo2NTM1NTE5NzAyMTA6a2V5LzhkMDk0MzNmLWFjZmUtNDI0NC1hMWI5LWU3MDk1YzhhZDU2ZVwifX1dLFwidGFnXCI6XCJZaE9kZk5UaWhheS1CRC1xZkVXNEpRXCJ9IiwicnMxIjoiMGQxZDdkODZhNGIxYzkzZWZlMjI4YWU1NGU0NGI3ZTVjOTM2Y2E5OCIsImF1ZCI6Ijg0MzE3ZTUyNTc0ZDRhOWJiNjU1NmNjOTI5NjgyZmYxZDA0NDY0OWEiLCJpc3MiOiJodHRwczovL2lkZW50aXR5LmNpYW0ucmVmaW5pdGl2LmNvbS9hcGkvaWRlbnRpdHkvc3RzX3ByZXByb2QiLCJleHAiOjE2NTI0NDY1MTMsImlhdCI6MTY1MjQ0NTkxM30.RAaoT4lrwk3JfpcHKQ4Gu7L4q8TA7UlsskEONTWpMfmikIwJvINRuzFEFvyBS45JtVI9g-Iw08IbeoZDewjis9cIg-7jjRsie2OEI5JhTR03XIQTHEhud9z6dcH4zVMeC1osaCSfyC5QeQUytnn_KgBs93QglC0720SjewjNq3kyUI3MK77QhfA95iWCn8-boguKGTAfcAHT_PlV_Jj4nrmrSlzJ07t14rOnQgunAiRr4QltBA00VKPyrLf1C9YwUr1aHyF-QAKMPp3ypDzGDeQNalcHMrtKHsyQZR8mvjMtht0AWFTff9kM-tcqIhbZQtJxRneAFNyGxAhVxbpOjw";
  console.log(token);
  const opts = {
    token: `Bearer ${token}`,
    timeout: -1,
  };
  session
    .open(opts)
    .then((val) => {
      console.log("data value: ", val);
    })
    .catch((e) => {
      console.error("data error", e);
    })
    .finally(() => {
      console.log("data finally");
    });
  session.on("open", () => {
    const subscription = session.createSubscription({
      name: "RDP.Data",
      args: ["CNY=;GBP=;CHF=;NZD=;SEK=", "CF_BID;CF_ASK;CF_NETCHNG", "CH=FD RH=IN", null],
    });
    subscription.on("update", (result: platform.FunctionResult) => {
      console.log(`function result: ${JSON.stringify(result)}`);
    });
    setInterval(() => subscription.refresh(), 1000);

    setInterval(() => session.updateToken(opts.token), 240000);
    setTimeout(() => {
      session.pause();
    }, 240000);
  });
}
