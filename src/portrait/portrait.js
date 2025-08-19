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
      // Récupérer la première section du document
      const firstSection = context.document.sections.getFirst();
  
      // 1. Obtenir l'objet pageSetup via la méthode getPageSetup()
      const pageSetup = firstSection.getPageSetup();
  
      // 2. Charger la propriété "orientation" sur l'objet pageSetup
      pageSetup.load("orientation");
  
      // Exécuter les commandes en attente (charger l'orientation)
      await context.sync();
  
      // 3. Accéder à la propriété maintenant qu'elle est chargée
      const orientation = pageSetup.orientation;
      console.log("Orientation : ", orientation);
  
      // Insérer un paragraphe avec l'orientation dans le document
      context.document.body.insertParagraph(
        "Orientation de la première section : " + orientation,
        Word.InsertLocation.end
      );
  
      await context.sync();
    });
  }