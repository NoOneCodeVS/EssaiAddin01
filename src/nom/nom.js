/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    document.getElementById("run").onclick = run;
});

export async function run() {
    await Word.run(async (context) => {
        const properties = context.document.properties;
        properties.load("author");
        await context.sync();
        const author = properties.author;
        console.log("Auteur du document :",  author);
        if(author === "Groupe Projet Alpha"){
            document.getElementById("error").style.display = "none";
            document.getElementById("resolve").style.display = "block";
            document.getElementById("exerice").style.display = "none";
        }
        else{
            document.getElementById("error").style.display = "block";
            document.getElementById("msg").innerText = `Le nom actuel est : ${author} il doit être : Groupe Projet Alpha`;
        }
    });
  }