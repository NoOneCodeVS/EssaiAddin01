/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    document.getElementById("nom_run").onclick = nom_run;
});

export async function nom_run() {
    await Word.run(async (context) => {
        const properties = context.document.properties;
        properties.load("author");
        await context.sync();
        const author = properties.author;
        console.log("Auteur du document :",  author);
        if(author === "Groupe Projet Alpha"){
            document.getElementById("nom_error").style.display = "none";
            document.getElementById("nom_resolve").style.display = "block";
            document.getElementById("nom_isResolve").checked = true;
        }
        else{
            document.getElementById("nom_error").style.display = "block";
            document.getElementById("nom_msg").innerText = `Le nom actuel est : ${author} il doit être : Groupe Projet Alpha`;
        }
    });
  }