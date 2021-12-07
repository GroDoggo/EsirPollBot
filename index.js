const { Client, Intents } = require('discord.js'); //CLIENT
const { MessageActionRow, MessageButton } = require('discord.js'); //BOUTTON
const client = new Client({ intents: [Intents.FLAGS.GUILDS, Intents.FLAGS.GUILD_MESSAGES] }); //LECTURE DES MESSAGES
const fs = require('fs'); //LECTURE DU TOKEN
var xl = require('excel4node');

var vote = {
  "1234567890": {
    "question": "",
    "author": "",
    "channel": "channelId",
    "no": [],
    "yes": [],
    "total": 0,
    "end": false,
    "max" : 0,
  }
}




//LOG ON START
client.on('ready', () => {
  console.log(`Logged in as ${client.user.tag}!`);
});

client.on('interactionCreate', async interaction => {

  if (interaction.isCommand()) {

    //VOTE
    if (interaction.commandName === 'vote') {

      const question = interaction.options.getString("question") //LA QUESTION POSE

      const resultat = "🔳🔳🔳🔳🔳🔳🔳🔳🔳🔳 (00%)" //ETAT ACTUEL DU VOTE

      //EMBED PAR DEFAUT
      var embed = {
        "title": "VOTE : " + question,
        "color": 12260372,
        "fields": [
          {
            "name": "Résultats :",
            "value": resultat,
            "inline": true
          }
        ],
        "author": {
          "name": "SEB"
        },
        "footer": {
          "text": "Super ESIR Bot"
        },
        "timestamp": new Date().toISOString()
      }

      //ENVOI DU MESSAGE
      interaction.channel.send({ embeds: [embed] })
        .then(msg => {

          //CREATION DES BOUTTONS
          const row = new MessageActionRow()
            .addComponents(
              new MessageButton()
                .setCustomId('yAns' + msg.id)
                .setLabel('OUI')
                .setStyle('PRIMARY'),
              new MessageButton()
                .setCustomId('nAns' + msg.id)
                .setLabel('NON')
                .setStyle('PRIMARY')
            );

          if (interaction.options.getBoolean("voteblanc")) row.addComponents(
            new MessageButton()
              .setCustomId('vBlc' + msg.id)
              .setLabel('BLANC')
              .setStyle('SECONDARY')
          )

          const end = new MessageActionRow()
            .addComponents(
              new MessageButton()
                .setCustomId('vEnd' + msg.id)
                .setLabel('METTRE FIN AU SONDAGE')
                .setStyle('DANGER'),
            );

          //Changement du maxVote si ce dernier a été paramétré
          var maxVote = -1
          if (interaction.options.getInteger("maxvote") != undefined) maxVote = interaction.options.getInteger("maxvote")

          //CREATION DU JSON
          vote[msg.id] = {
            "question": question.replace(/[^a-zA-Z ]/gi, ""),
            "author": interaction.user.id,
            "channel": msg.channel.id,
            "no": [],
            "yes": [],
            "vBlc" : [],
            "total": 0,
            "end": false,
            "max" : maxVote
          }

          //ENVOI DES BOUTONS
          msg.edit({ components: [row] })
          interaction.reply({ ephemeral: true, content: "Votre commande a bien été envoyée" })
          interaction.user.send({ content: "Ton sondage [" + question + "] a bien été crée, tu peux l'arreter a tout moment avec ce bouton", components: [end] })
        })
    }
  }

  if (interaction.isButton()) {

    const buttonId = interaction.customId
    const voteId = buttonId.slice(4)

    if (vote[voteId] != undefined) {

      if (buttonId === ("yAns" + voteId)) { //Le bouton cliclé est OUI

        if (vote[voteId]["end"]) return;

        const length = vote[voteId]["no"].length
        vote[voteId]["yes"] = vote[voteId]["yes"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU OUI
        vote[voteId]["no"] = vote[voteId]["no"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU NON
        vote[voteId]["vBlc"] = vote[voteId]["vBlc"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU VOTE BLANC
        vote[voteId]["yes"].push(interaction.user.id) //ON AJOUTE DANS LE YES
        vote[voteId]["total"] =  vote[voteId]["yes"].length + vote[voteId]["no"].length + vote[voteId]["vBlc"].length //Calcul du total
        if (vote[voteId]["max"] > 0 && vote[voteId]["total"] >= vote[voteId]["max"]) vote[voteId]["end"] = true; //Si le max est atteint => STOP

      } else if (buttonId === ("nAns" + voteId)) { //Le bouton cliqué est NON

        if (vote[voteId]["end"]) return;

        const length = vote[voteId]["yes"].length
        vote[voteId]["yes"] = vote[voteId]["yes"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU OUI
        vote[voteId]["no"] = vote[voteId]["no"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU NON
        vote[voteId]["vBlc"] = vote[voteId]["vBlc"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU VOTE BLANC
        vote[voteId]["no"].push(interaction.user.id) //ON AJOUTE DANS LE NON
        vote[voteId]["total"] =  vote[voteId]["yes"].length + vote[voteId]["no"].length + vote[voteId]["vBlc"].length //Calcul du total
        if (vote[voteId]["max"] > 0 && vote[voteId]["total"] >= vote[voteId]["max"]) vote[voteId]["end"] = true; //Si le max est atteint => STOP

      } else if (buttonId === ("vBlc" + voteId)) { //Le bouton cliqué est VOTE BLANC

        if (vote[voteId]["end"]) return;

        const length = vote[voteId]["yes"].length
        vote[voteId]["yes"] = vote[voteId]["yes"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU OUI
        vote[voteId]["no"] = vote[voteId]["no"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU NON
        vote[voteId]["vBlc"] = vote[voteId]["vBlc"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU VOTE BLANC
        vote[voteId]["vBlc"].push(interaction.user.id) //ON AJOUTE DANS LE NON
        vote[voteId]["total"] =  vote[voteId]["yes"].length + vote[voteId]["no"].length + vote[voteId]["vBlc"].length //Calcul du total
        if (vote[voteId]["max"] > 0 && vote[voteId]["total"] >= vote[voteId]["max"]) vote[voteId]["end"] = true; //Si le max est atteint => STOP

      } else if (buttonId === ("vEnd" + voteId)) { //Le bouton cliqué est FINIR LE VOTE

        if (vote[voteId]["end"]) { //Si le vote est déja terminé
          interaction.reply("Ce vote semble etre déja terminé...")
          return;
        }
        createExcel(vote[voteId]) //On crée l'excel
        interaction.reply("Le vote est terminé, tu va recevoir le resultat...") //On informe l'utilisateur
        //ATTENTE ET ENVOI DU EXCEL
        setTimeout(() => {
          console.log("SENDING...")
          client.users.fetch(vote[voteId]["author"])
            .then(u => {
              u.send({ files: ["./Sondage/Sondage_" + vote[voteId]["question"].replace(/ /gi, '_') + ".xlsx"] })
            })
        }, 5000)

        vote[voteId]["end"] = true;

      } else return;

      //FETCHING CHANNEL
      client.channels.fetch(vote[voteId]["channel"])
        .then(channel => {
          if (channel.isText()) {

            //FETCHING MESSAGE
            channel.messages.fetch(voteId)
              .then(message => {
                if (message.embeds.length == 1 && message.embeds[0].fields.length == 1) {
                  if (!vote[voteId]["end"]) { //LE VOTE NEST PAS TERMINER
                    const resultat = getResultat(vote[voteId]["yes"].length, vote[voteId]["no"].length)
                    const modif = message.embeds[0]
                    modif.setFields([
                      {
                        "name": "Résultats :",
                        "value": resultat,
                        "inline": true
                      }
                    ]
                    )

                    //MOFICATION DU MESSAGE
                    message.edit({ embeds: [modif] })
                    interaction.reply({ ephemeral: true, content: "Ton vote a bien été pris en compte\nTu peux le modifier en cliquant de nouveau sur un choix" })
                  } else { //LE VOTE EST TERMINER
                    const resultat = getResultat(vote[voteId]["yes"].length, vote[voteId]["no"].length)
                    var resultatFinal = "OUI"
                    if (vote[voteId]["yes"].length < vote[voteId]["no"].length) var resultatFinal = "NON"
                    const modif = message.embeds[0]
                    modif.setFields([
                      {
                        "name": "Résultats :",
                        "value": resultat,
                        "inline": false
                      }
                    ],
                    [
                      {
                        "name": "Ce vote est terminé",
                        "value": "Le resultat final est " + resultatFinal,
                        "inline": false
                      }
                    ])
                    //MOFICATION DU MESSAGE
                    message.edit({ embeds: [modif] })
                  }
                }
              })

          }
        })
    }

  }

})

//CALCUL DU POURCENTAGE ET CREATION DE LA BARRE
function getResultat(positif, negatif) {

  var res = ""
  const pourcentage = Math.round(positif / (positif + negatif) * 100)
  for (var i = 0; i < pourcentage / 10; i++) res += "⬜"
  for (var i = 10; i > pourcentage / 10; i--) res += "🔳"

  return res + " (" + pourcentage + "%)"
}

//CREATION DE L'EXCEL
function createExcel(stats) {

  //INITIALISATION DU EXCEL
  var wb = new xl.Workbook();
  var ws = wb.addWorksheet('STATS SONDAGE');
  const question = stats["question"].replace(/ /gi, '_');
  var style = wb.createStyle({
    font: {
      color: '#000000',
      size: 12,
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -',
  });

  const colOUI = 3
  const colNON = 4

  ws.cell(2, colOUI)
    .string("REPONSE POSITIVE")
    .style(style)
    .style({ font: { size: 14 } });

  ws.cell(2, colNON)
    .string("REPONSE NEGATIVE")
    .style(style)
    .style({ font: { size: 14 } });

  for (var i = 0; i < stats["yes"].length; i++) {
    client.users.fetch(stats["yes"][i])
      .then(u => {
        ws.cell(i + 3, colOUI)
          .string(u.username)
          .style(style);
      })
  }

  for (var i = 0; i < stats["no"].length; i++) {
    client.users.fetch(stats["no"][i])
      .then(u => {
        ws.cell(i + 3, colNON)
          .string(u.username)
          .style(style);
      })
  }

  ws.column(colOUI).setWidth(50);
  ws.column(colNON).setWidth(50);

  wb.write("./Sondage/Sondage_" + question + ".xlsx")
}

client.login(fs.readFileSync('token.txt', 'utf8'));
