const { Client, Intents } = require('discord.js'); //CLIENT
const { MessageActionRow, MessageButton } = require('discord.js'); //BOUTTON
const client = new Client({ intents: [Intents.FLAGS.GUILDS, Intents.FLAGS.GUILD_MESSAGES] }); //LECTURE DES MESSAGES
const fs = require('fs'); //LECTURE DU TOKEN
var xl = require('excel4node');

var vote = {
  "1234567890": {
    "question" : "",
    "author" : "",
    "channel": "channelId",
    "no": [],
    "yes": [],
    "total": 0
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

          //CREATION DU JSON
          vote[msg.id] = {
            "question" : question,
            "author" : interaction.user.id,
            "channel": msg.channel.id,
            "no": [],
            "yes": [],
            "total": 0
          }

          //ENVOI DES BOUTONS
          msg.edit({ components: [row] })
          interaction.reply({ephemeral : true, content : "Votre message a bien été envoyé"})
          interaction.user.send({content : "Ton sondage [" + question + "] a bien été crée, tu peux l'arreter a tout moment avec ce bouton", components : [end]})
        })
    }
  }

  if (interaction.isButton()) {

    const buttonId = interaction.customId
    const voteId = buttonId.slice(4)

    if (vote[voteId] != undefined) {

      if (buttonId === ("yAns" + voteId)) { //Le bouton cliclé est OUI

        const length = vote[voteId]["no"].length
        vote[voteId]["yes"] = vote[voteId]["yes"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU OUI
        vote[voteId]["no"] = vote[voteId]["no"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU NON
        vote[voteId]["yes"].push(interaction.user.id) //ON AJOUTE DANS LE YES
        if (vote[voteId]["no"].length == length) vote[voteId]["total"] += 1 //Si aucun changement de vote on ajoute 1

      } else if (buttonId === ("nAns" + voteId)) { //Le bouton cliqué est NON

        const length = vote[voteId]["yes"].length
        vote[voteId]["yes"] = vote[voteId]["yes"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU OUI
        vote[voteId]["no"] = vote[voteId]["no"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU NON
        vote[voteId]["no"].push(interaction.user.id) //ON AJOUTE DANS LE NON
        if (vote[voteId]["yes"].length == length) vote[voteId]["total"] += 1 //Si aucun changement de vote on ajoute 1

      } else if (buttonId === ("vBlc" + voteId)){ //Le bouton cliqué est VOTE BLANC

        const length = vote[voteId]["yes"].length
        vote[voteId]["yes"] = vote[voteId]["yes"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU OUI
        vote[voteId]["no"] = vote[voteId]["no"].filter(u => !(u === interaction.user.id)) //ON SUPPRIME DU NON
        if (vote[voteId]["yes"].length == length) vote[voteId]["total"] += 1 //Si aucun changement de vote on ajoute 1

      }else if (buttonId === ("vEnd" + voteId)){ //Le bouton cliqué est VOTE BLANC

        createExcel(vote[voteId])
        setTimeout(() => {
          console.log("SENDING...")
          client.users.fetch(vote[voteId]["author"])
          .then(u => {
            u.send({files : ["./Sondage/Sondage " + vote[voteId]["question"] + ".xlsx"]})
          }) 
        }, 5000)

        return;

      } else return;

      //FETCHING CHANNEL
      client.channels.fetch(vote[voteId]["channel"])
        .then(channel => {
          if (channel.isText()) {

            //FETCHING MESSAGE
            channel.messages.fetch(voteId)
              .then(message => {
                if (message.embeds.length == 1 && message.embeds[0].fields.length == 1) {
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
                  interaction.reply({ephemeral : true, content : "Ton vote a bien été pris en compte\nTu peux le modifier en cliquant de nouveau sur un choix"})
                }
              })

          }
        })
    }

  }

})

function getResultat(positif, negatif) {

  var res = ""
  const pourcentage = Math.round(positif / (positif + negatif) * 100)
  for (var i = 0; i < pourcentage / 10; i++) res += "⬜"
  for (var i = 10; i > pourcentage / 10; i--) res += "🔳"

  return res + " (" + pourcentage + "%)"
}

function createExcel(stats){

  //INITIALISATION DU EXCEL
  var wb = new xl.Workbook();
  var ws = wb.addWorksheet('STATS SONDAGE');
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
  .style({font: {size: 14}});

  ws.cell(2, colNON)
  .string("REPONSE NEGATIVE")
  .style(style)
  .style({font: {size: 14}});

  for (var i = 0; i < stats["yes"].length; i++){
    client.users.fetch(stats["yes"][i])
    .then (u => {
      ws.cell(i+3, colOUI)
      .string(u.username)
      .style(style);
    })    
  }

  for (var i = 0; i < stats["no"].length; i++){
    client.users.fetch(stats["no"][i])
    .then (u => {
      ws.cell(i+3, colNON)
      .string(u.username)
      .style(style);
    })
  }

  ws.column(colOUI).setWidth(50);
  ws.column(colNON).setWidth(50);

  wb.write("./Sondage/Sondage " + stats["question"] + ".xlsx")
}

client.login(fs.readFileSync('token.txt', 'utf8'));
