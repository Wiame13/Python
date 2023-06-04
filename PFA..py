import pandas as pd
import re

#1.recuperer le fichier excel.
xl = pd.read_excel(r"C:\Users\lenovo\Desktop\data11.xlsx")
xl['Participants_ID'] = xl.index
#recuperer chaque tableau du fichier excel
profil =pd.DataFrame(xl[['Participants_ID','Horodateur','Genre','Âge','Situation familiale','Situation professionnelle',"Niveau d'étude",'Ville']])
qst1_1x =pd.DataFrame(xl[['Participants_ID',"Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"]])

don = pd.DataFrame(xl[['Participants_ID',"Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?",
                       'Question 1.2 - Que pensez-vous de la situation du sang dans votre pays?',"Question 1.3 - Avez-vous déjà eu un besoin urgent de sang ou l'un de vos proches ? (Jamais (0) jusqu’à souvent (5)",
                      'Question 2- Connaissez-vous votre groupe sanguin ?','Question 3- Avez-vous déjà fait au moins une fois don de votre sang ?']])
motiv = pd.DataFrame(xl[['Participants_ID','Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Volontariat pour aider les patients ]',
                         "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Une façon pour moi de m'intégrer dans la communauté et d'y contribuer]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Donner du sang fait partie de mes obligations ]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Les convictions religieuses]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important))) [La bonne communication et le service des centres de don de sang]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [La confiance dans les centres de don du sang ]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Bilan de santé et groupe sanguin gratuit ]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Cadeau, récompenses financières, billets ]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Demande de vos proches (famille) ou un ami a besoin de sang ]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Curiosité sur le don de sang ]",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important))[Mon groupe sanguin est rare et recherché ]",
                        "Question 5 - Avez-vous déjà incité vos proches à faire don de leur sang ? (Jamais (0) jusqu’à souvent (5)"]])
motiv.rename(columns={'Participants_ID':'Participants_ID','Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Volontariat pour aider les patients ]':"Volontariat pour aider les patients",
                         "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Une façon pour moi de m'intégrer dans la communauté et d'y contribuer]":"Une façon pour moi de m'intégrer dans la communauté et d'y contribuer",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Donner du sang fait partie de mes obligations ]":"Donner du sang fait partie de mes obligations",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Les convictions religieuses]":"Les convictions religieuses",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important))) [La bonne communication et le service des centres de don de sang]":"La bonne communication et le service des centres de don de sang",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [La confiance dans les centres de don du sang ]":"La confiance dans les centres de don du sang ",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Bilan de santé et groupe sanguin gratuit ]":"Bilan de santé et groupe sanguin gratuit",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Cadeau, récompenses financières, billets ]":"Cadeau, récompenses financières, billets ",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Demande de vos proches (famille) ou un ami a besoin de sang ]":"Demande de vos proches (famille) ou un ami a besoin de sang",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important)) [Curiosité sur le don de sang ]":"Curiosité sur le don de sang",
                        "Question 4 - Quels sont les motivations qui vous poussent à donner du sang?((Pas du tout important) à (extrêmement important))[Mon groupe sanguin est rare et recherché ]":"Mon groupe sanguin est rare et recherché",
                        "Question 5 - Avez-vous déjà incité vos proches à faire don de leur sang ? (Jamais (0) jusqu’à souvent (5)":"Question 5 - Avez-vous déjà incité vos proches à faire don de leur sang ? (Jamais (0) jusqu’à souvent (5)"
},inplace=True)
qst_6 = pd.DataFrame(xl[['Participants_ID',"Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Peur des aiguilles, vue du sang ]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [ Peur d'être infecté par le virus covid19]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Effets indésirables : Fièvre, Perte de poids, Hypertension artérielle, Vertiges, ]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Douloureux / dangereux]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Inconvenable lieux et heures d'ouverture des centres des dons ]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Manque de places de parking]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Contrainte de temps : être trop occupé (travail, famille)]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Je ne peux pas faire de don en raison de mon état de santé ]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Apathie, jamais pensé à donner du sang]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Manque des connaissances du processus de don de sang ]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Manque de conscience du besoin de sang]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Le sang est donné gratuitement et ensuite vendu ]",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Pas de confiance aux centres de don du sang ]"]])
qst_6.rename(columns={"Participants_ID":"Participants_ID","Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Peur des aiguilles, vue du sang ]":"Peur des aiguilles, vue du sang",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [ Peur d'être infecté par le virus covid19]":"Peur d'être infecté par le virus covid19",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Effets indésirables : Fièvre, Perte de poids, Hypertension artérielle, Vertiges, ]":"Effets indésirables : Fièvre, Perte de poids, Hypertension artérielle, Vertiges",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Douloureux / dangereux]":"Douloureux / dangereux",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Inconvenable lieux et heures d'ouverture des centres des dons ]":"Inconvenable lieux et heures d'ouverture des centres des dons",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Manque de places de parking]":"Manque de places de parking",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Contrainte de temps : être trop occupé (travail, famille)]":"Contrainte de temps : être trop occupé (travail, famille)",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Je ne peux pas faire de don en raison de mon état de santé ]":'Je ne peux pas faire de don en raison de mon état de santé',
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Apathie, jamais pensé à donner du sang]":"Apathie, jamais pensé à donner du sang",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Manque des connaissances du processus de don de sang ]":"Manque des connaissances du processus de don de sang",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important)) [Manque de conscience du besoin de sang]":"Manque de conscience du besoin de sang",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Le sang est donné gratuitement et ensuite vendu ]":"Le sang est donné gratuitement et ensuite vendu",
                        "Question 6 - Comment qualifieriez-vous les peurs et les idées fausses empêchant les donneurs de faire un don ? ((Extrêmement important) à (Pas du tout important))[Pas de confiance aux centres de don du sang ]":"Pas de confiance aux centres de don du sang"},inplace=True)
incit = pd.DataFrame(xl[['Participants_ID',"Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Certificat, carte de fidélité ]",
                         "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé))[Cadeaux (parapluie, radio, tasse, porte-clés, lampe de poche, T-shirt,chapeau..) ]",
                         "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Récompenses financières (coupons gratuits, parking, ticket) ]",
                        "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Connaître le groupe sanguin ]",
                        "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Avoir un bilan de santé ]",
                        "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Appréciation des donateurs via les médias (journaux, site Web, radio, télévision) ]"]])
incit.rename(columns={"Participants_ID":"Participants_ID","Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Certificat, carte de fidélité ]":"Certificat, carte de fidélité",
                         "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé))[Cadeaux (parapluie, radio, tasse, porte-clés, lampe de poche, T-shirt,chapeau..) ]":"Cadeaux (parapluie, radio, tasse, porte-clés, lampe de poche, T-shirt,chapeau..)",
                         "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Récompenses financières (coupons gratuits, parking, ticket) ]":"Récompenses financières (coupons gratuits, parking, ticket)",
                        "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Connaître le groupe sanguin ]":"Connaître le groupe sanguin",
                        "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Avoir un bilan de santé ]":"Avoir un bilan de santé",
                        "Question 7 - Parmi les incitations suivantes, lesquelles souhaiteriez-vous recevoir lorsque vous donnez du sang? ((Extrêmement intéressé) à (Pas du tout intéressé)) [Appréciation des donateurs via les médias (journaux, site Web, radio, télévision) ]":"Appréciation des donateurs via les médias (journaux, site Web, radio, télévision)"},inplace=True)
com = pd.DataFrame(xl[["Participants_ID","Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Via une application mobile (qui contient des informations importantes pour donner et recevoir du sang)]",
                       "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Des courriers (mail), des SMS )]",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Les réseaux sociaux ]",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Les médias (TV, radio…) ]",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [La presse locale, nationale ]",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Présence dans: événements, grandes surfaces, foires, souk hebdomadaire ]",
                       "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important))[Via d'autre solutions innovantes de gestion de donneurs (les rendez-vous, gestion et traçabilité..) ]",
                      "Question 9 -Combien de fois avez-vous l'intention de donner du sang l'année prochaine ?",
                      "Question 10 - Quelles sont vos suggestions pour améliorer la gestion et la promotion du don du sang?"]])
com.rename(columns={"Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Via une application mobile (qui contient des informations importantes pour donner et recevoir du sang)]":"Via une application mobile (qui contient des informations importantes pour donner et recevoir du sang)",
                       "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Des courriers (mail), des SMS )]":"Des courriers (mail), des SMS )",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Les réseaux sociaux ]":"Les réseaux sociaux",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Les médias (TV, radio…) ]":"Les médias (TV, radio…)",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [La presse locale, nationale ]":"La presse locale, nationale",
                      "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important)) [Présence dans: événements, grandes surfaces, foires, souk hebdomadaire ]":"Présence dans: événements, grandes surfaces, foires, souk hebdomadaire",
                       "Question 8 - Quelles sont les sources que vous préférez pour recevoir et transmettre les demandes de dons de sang ou les publicités de campagne ((Extrêmement important) à (Pas du tout important))[Via d'autre solutions innovantes de gestion de donneurs (les rendez-vous, gestion et traçabilité..) ]":"Via d'autre solutions innovantes de gestion de donneurs (les rendez-vous, gestion et traçabilité..)",
                      "Question 9 -Combien de fois avez-vous l'intention de donner du sang l'année prochaine ?":"Question 9 - Quelle est la probabilité que vous fassiez des dons de sang dans le futur ?"},inplace=True)

mngmtx= pd.DataFrame(xl.loc[xl['Question 3- Avez-vous déjà fait au moins une fois don de votre sang ?'] == 'OUI  نعم', ['Participants_ID', 'Question 3.1- Si oui, combien de fois avez-vous fait de don ?','Question 3.2 - Où effectuez-vous généralement votre don de sang ?','Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)']])
qst3_4= pd.DataFrame(xl.loc[xl['Question 3- Avez-vous déjà fait au moins une fois don de votre sang ?'] == 'OUI  نعم',['Participants_ID','Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Accueil]',"Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Salle d'attente avant le don]",
                        'Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Durant le don ]',"Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Après le don ]"]])
qst3_4.rename(columns={'Participants_ID': 'Participants_ID',
                       "Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Accueil]": 'Satisfaction_Accueil',
                       "Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Salle d'attente avant le don]": 'Satisfaction_Salle_Attente',
                       "Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Durant le don ]": 'Satisfaction_Pendant_Don',
                       "Question 3.4 - Quel est votre degré de satisfaction du processus de don de sang ? [Après le don ]": 'Satisfaction_Apres_Don'},
              inplace=True)

exp =pd.DataFrame(xl.loc[xl['Question 3- Avez-vous déjà fait au moins une fois don de votre sang ?'] == 'OUI  نعم',['Participants_ID',"Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important)) [J'ai beaucoup appris et bénéficié (la valeur du don, le groupe sanguin, les informations sur le don et le sang..)]",
                      "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important))[Je pense que je donnerai plus souvent ]",
                     "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important))[C’est très différent des dons que j’ai pu faire à d'autres associations (argent, nourriture....) ]",
                     "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important)) [C’est devenu une habitude ]",
                     "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important))[Je me sens de plus en plus à l’aise lorsque je vais donner mon sang ]"]])
exp.rename(columns={'Participants_ID': 'Participants_ID',
                    "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important)) [J'ai beaucoup appris et bénéficié (la valeur du don, le groupe sanguin, les informations sur le don et le sang..)]": 'J\'ai beaucoup appris et bénéficié',
                    "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important))[Je pense que je donnerai plus souvent ]": 'Je pense que je donnerai plus souvent ',
                    "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important))[C’est très différent des dons que j’ai pu faire à d'autres associations (argent, nourriture....) ]": 'C’est très différent des dons que j’ai pu faire à d\'autres associations',
                    "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important)) [C’est devenu une habitude ]": 'C’est devenu une habitude',
                    "Question 3.5 - Quelle expérience ou quelle leçon avez-vous tiré des dons de sang que vous avez déjà faits ?((Pas du tout important) à (extrêmement important))[Je me sens de plus en plus à l’aise lorsque je vais donner mon sang ]": 'Je me sens de plus en plus à l’aise lorsque je vais donner mon sang'},
           inplace=True)




#2.les modification:
# choix multiple dans situation profitionnel
def remove_multi_choix(text):
    if isinstance(text, str):
        if ',' in text :
            if ' Vous exercez un travail ou une activité' in text :
                return "Travail/activité"
            if 'Etudiant' in text :
                return 'Etudiant'
        return text
    else : 
        return text
     # Définir une fonction pour supprimer les termes en arabe
def remove_arabic(text):
    if isinstance(text, str):
        return ''.join(c for c in text if not ('\u0600' <= c <= '\u06FF'))
    else:
        return text
def remove_parentheses(text):
    return re.sub(r'\([^()]*\)', '', text)
      #profil
profil['Genre']=profil['Genre'].apply(remove_arabic).str.strip()
profil['Âge']= profil['Âge'].apply(remove_arabic).str.strip()
profil['Situation familiale']= profil['Situation familiale'].apply(remove_arabic).apply(remove_parentheses).str.strip()
profil['Situation professionnelle']= profil['Situation professionnelle'].apply(remove_arabic).apply(remove_parentheses).str.strip().apply(remove_multi_choix).replace({"Vous exercez un travail ou une activité":"Travail/activité","Retraité":"Retraité(e)"})
                  #le modification sur niveau d'etude:
replacment={'رأي آخر Autre':"Autre","مستوى البكالوريا                      Niveau bac":"Niveau d'étude","البكالوريا ->  البكالوريا+2                Bac -> Bac+2":"Bac-Bac+2","Bac+2-> Bac+3            البكالوريا+2 -> البكالوريا+3":"Bac+2-Bac+3","البكالوريا+3 ->  البكالوريا+5           Bac+3 -> Bac+5":"Bac+3- Bac+5","وما فوق البكالوريا+Bac+5              5 >":"Bac+ 5 >"}
profil["Niveau d'étude"]=profil["Niveau d'étude"].replace(replacment)
#cleaning de ville 
cities= {"A":'','Agadir':['Agadir','AGADIR','Agadir','agadir'],'Azrou':['Azrou','azrou'],'Béni Mellal':['Benimellal','Beni Mellal','بني ملال','zaouiat cheikh','Béni Mellal','Beni MELLAL','BENI MELLAL','Beni mellal','Béni mellal','beni mellal'],
         'Berrechid':['Berchid','BERCHID'],
    'Casablanca': ['Casablanca', 'Casa Blanca', 'Casa', 'casa','Casablanca-Rabat','casablanca'],'Chefchaouen':['chefchaouen'],'Driouch':['DRIOUCH'],
     'Dar ould zidouh':['Dar ould zidouh','دار ولد زيدوح'],'El Jadida':['El Jadida','El jadida','el jadida','El Jadida ','EL JADIDA','الجديدة','jadida'],'El Kelâa des Sraghna':['el kalaa des sraghna','El Kelâa des Sraghna','Kelaa des Sraghna'],
     'Errachidia':['errachidia','Errachidia','الرشيدية'] ,'Fès':['Fès','فاس','FES','Fes','Fés','fes'],"M'diq":["M'diq","المضيق"],
        'Fquih Ben Salah':['Fquih Ben Salah','Fquih ben salah','fkih ben saleh','Fkih Ben salah','Fkih Ben Salah','Fquih ben salah','الفقيه بن صالح','فقيه بن صالح'],'Safi':['Safi','SAFI'],
    'Rabat': ['RABAT','Rabat', 'Harhoura','Rabat agadir','rabat','الرباط'],'Kénitra':['Kénitra','kenitra','Kenitra',],'Khouribga':['Khouribga','khouribgua','KHOURIBGUA',],
        'Ksar El-Kébir':['Ksar El-Kébir','ksar el kebir','Ksar el Kebir'],'Laâyoune':['Laâyoune','laayoune','Laayoune'],'Larache':['Larache','larache'],
         'Marrakech': ['Marrakech', 'Marrakesh','MARRAKECH','مراكش'],'Meknès':['Meknès','Meknes','meknes','MEKNES'],'Oujda':['Oujda','oujda'],'Mohammedia':['Mohammedia','محمديه','mohammedia'],
      'Sala al jadida':['Sala al jadida','Sala eljadida'],'Salé':['Salé','سلا','sale','Sale','SALE','salé'],
         'Settat':['settat','Settat'],'Skhirat':['Skhirat','skhirate'],'Tadla':['tadla','Tadla'],'Tanger':['TANGER','Tanger','tanger'],
         'Témara':['témara','temara','Temera','Témara','Temara'],'Tétouan':['Tétouan','tetouan','Tetouan','تطوان'],'Tiznit':['Tiznit','tiznit']
         
}
for city, variations in cities.items():
    profil['Ville'] = profil['Ville'].replace(variations, city).str.strip()

#creation d'un colonne des regions 
correspondance_villes_regions = {
    'Tanger-Tétouan-Al Hoceïma': ['Ksar El-Kébir','Martil','Tanger', 'Tétouan', 'Al Hoceïma', "M\'diq", 'Larache', 'Chefchaouen', 'Ouezzane'],
    'Oriental': ['Oujda', 'Nador', 'Berkane', 'Taourirt', 'Jerada', 'Driouch'],
    'Fès-Meknès': ['Fès', 'Meknès','Azrou','Sefrou', 'Taza', 'Ifrane', 'Boulemane'],
    'Rabat-Salé-Kénitra': ['Rabat','Sala al jadida','Salé', 'Kénitra', 'Skhirat','Témara', 'Khémisset','Sidi slimane'],
    'Casablanca-Settat': ['Casablanca','Bouznika', 'Mohammedia', 'Settat', 'El Jadida', 'Nouaceur', 'Berrechid'],
    'Marrakech-Safi': ['Marrakech', 'Essaouira', 'Safi', 'Chichaoua', 'Al Haouz', 'El Kelâa des Sraghna', 'Youssoufia'],
    'Souss-Massa': ['Agadir', 'Inezgane', 'Tiznit', 'Taroudant', 'Chtouka-Aït Baha'],
    'Béni Mellal-Khénifra':['Béni Mellal','Dar ould zidouh','Khenifra','Khouribga','Azilal','Tadla','Fquih Ben Salah'],
    'Drâa Tafilalet':['Errachidia','Ouarzazate','Midelt','Zagora','Er-rich'],
    'Laâyoune-Sakia El Hamra':['Laâyoune','Boujdour','Tarfaya'],
    'Guelmim-Oued Noun':['Guelmim','Sidi Ifni','Tan-Tan']
}           
profil['Region']=profil['Ville']
for reg, var in correspondance_villes_regions.items():
    profil['Region'] = profil['Region'].replace(var, reg) 
      #don 

don["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"] = don["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"].apply(remove_arabic).apply(remove_parentheses).str.replace('(','').str.replace('  ','').str.replace('Communication à travers internet et les réseaux sociaux','Internet et les réseaux sociaux')
don["Question 1.2 - Que pensez-vous de la situation du sang dans votre pays?"] = don["Question 1.2 - Que pensez-vous de la situation du sang dans votre pays?"].apply(remove_arabic).apply(remove_parentheses).str.replace('(','').str.replace('  ','').replace({"Il y a un déficit de sang":"Déficit","Il y a un équilibre entre demande et offre de sang ":"Equilibre","Il y a un excédent de sang":"Excédent"})   
don["Question 2- Connaissez-vous votre groupe sanguin ?"]=don["Question 2- Connaissez-vous votre groupe sanguin ?"].apply(remove_arabic).str.strip().replace({'OUI':'Oui','NON':'Non'})
don["Question 3- Avez-vous déjà fait au moins une fois don de votre sang ?"]=don["Question 3- Avez-vous déjà fait au moins une fois don de votre sang ?"].apply(remove_arabic).str.strip().replace({'OUI':'Donneurs','NON':'Non Donneurs'})
don["Question 1.3 - Avez-vous déjà eu un besoin urgent de sang ou l'un de vos proches ? (Jamais (0) jusqu’à souvent (5)"]=don["Question 1.3 - Avez-vous déjà eu un besoin urgent de sang ou l'un de vos proches ? (Jamais (0) jusqu’à souvent (5)"].replace({0:"Jamais",1:"Rarement",2:"Parfois",3:"Parfois",4:"Souvent",5:"Plus souvent"})    
    #question 1.1
             # Appliquer la fonction sur la colonne

qst1_1x["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"] = qst1_1x["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"].apply(remove_arabic).apply(remove_parentheses).str.replace('(','').str.replace('  ','').str.replace(' Appel aux dons',' Appel télphonique aux dons').str.replace('Communication à travers internet et les réseaux sociaux','Internet et les réseaux sociaux')
             #duplication des ligne au cas ou on a plus d'une reponse dans une ligne
new_rows = []
for index, row in qst1_1x.iterrows():
    if ',' in row["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"]:
        values = row["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"].split(',')
        for value in values:
            new_rows.append({'Participants_ID': row['Participants_ID'], "Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?": value})
    else:
        new_rows.append({'Participants_ID': row['Participants_ID'], "Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?": row["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"]})
qst1_1 =pd.DataFrame(new_rows)
qst1_1["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"]=qst1_1["Question 1.1- Comment avez-vous entendu parler d'événement de don de sang ?"].str.strip()

    #mngmnt des donneures
new_rows2 = []
for index, row in mngmtx.iterrows():
    if ',' in row["Question 3.2 - Où effectuez-vous généralement votre don de sang ?"]:
        values = row["Question 3.2 - Où effectuez-vous généralement votre don de sang ?"].split(',')
        for value in values:
            new_rows2.append({'Participants_ID': row['Participants_ID'],'Question 3.1- Si oui, combien de fois avez-vous fait de don ?':row['Question 3.1- Si oui, combien de fois avez-vous fait de don ?'], "Question 3.2 - Où effectuez-vous généralement votre don de sang ?": value})
    else:
        new_rows2.append({'Participants_ID': row['Participants_ID'],'Question 3.1- Si oui, combien de fois avez-vous fait de don ?':row['Question 3.1- Si oui, combien de fois avez-vous fait de don ?'], "Question 3.2 - Où effectuez-vous généralement votre don de sang ?": row["Question 3.2 - Où effectuez-vous généralement votre don de sang ?"]})
mngmt=pd.DataFrame(new_rows2)
mngmt["Question 3.2 - Où effectuez-vous généralement votre don de sang ?"]=mngmt["Question 3.2 - Où effectuez-vous généralement votre don de sang ?"].apply(remove_arabic).str.replace(r'\(\s*\)',"").str.replace(r"\(D","D").str.strip()
mngmt["Question 3.2 - Où effectuez-vous généralement votre don de sang ?"]=mngmt["Question 3.2 - Où effectuez-vous généralement votre don de sang ?"].replace("Dans les collectes fixes (centres de don)","Dans les collectes fixes").replace("Dans les collectes fixes (centres des dons)","Dans les collectes fixes")
 
    #mnagment  2
new_rows3 = []    
for index, row in mngmtx.iterrows():
    if ',' in row["Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)"]:
        values = row["Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)"].split(',')
        for value in values:
            new_rows3.append({'Participants_ID': row['Participants_ID'], "Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)": value})
    else:
        new_rows3.append({'Participants_ID': row['Participants_ID'], "Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)": row["Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)"]})
mngmt2=pd.DataFrame(new_rows3)
mngmt2["Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)"]=mngmt2["Question 3.3 - Quels types de dons de sang faîtes-vous ? (Plusieurs réponses possibles)"].apply(remove_arabic).str.replace('  ','')
    #motiv 
columns=["Volontariat pour aider les patients",
"Une façon pour moi de m'intégrer dans la communauté et d'y contribuer",
"Donner du sang fait partie de mes obligations",
"Les convictions religieuses",
"La bonne communication et le service des centres de don de sang",
"La confiance dans les centres de don du sang ",
"Bilan de santé et groupe sanguin gratuit",
"Cadeau, récompenses financières, billets ",
"Demande de vos proches (famille) ou un ami a besoin de sang",
"Curiosité sur le don de sang",
"Mon groupe sanguin est rare et recherché"]

for col in columns :
    motiv[col]=motiv[col].apply(remove_arabic).str.strip()
motiv["Question 5 - Avez-vous déjà incité vos proches à faire don de leur sang ? (Jamais (0) jusqu’à souvent (5)"]=motiv["Question 5 - Avez-vous déjà incité vos proches à faire don de leur sang ? (Jamais (0) jusqu’à souvent (5)"].replace({0:"Jamais",1:"Rarement",2:"Parfois",3:"Souvent",4:"Plus souvent",5:"Toujours"})
    #qst 6
columnsqst6=["Peur des aiguilles, vue du sang",
"Peur d'être infecté par le virus covid19",
"Effets indésirables : Fièvre, Perte de poids, Hypertension artérielle, Vertiges",
"Douloureux / dangereux",
"Inconvenable lieux et heures d'ouverture des centres des dons",
"Manque de places de parking",
"Contrainte de temps : être trop occupé (travail, famille)",
'Je ne peux pas faire de don en raison de mon état de santé',
"Apathie, jamais pensé à donner du sang",
"Manque des connaissances du processus de don de sang",
"Manque de conscience du besoin de sang",
"Le sang est donné gratuitement et ensuite vendu",
"Pas de confiance aux centres de don du sang"]
for coll in columnsqst6 :
    qst_6[coll]=qst_6[coll].apply(remove_arabic).str.strip()
    #incit
columnsincit=["Certificat, carte de fidélité","Cadeaux (parapluie, radio, tasse, porte-clés, lampe de poche, T-shirt,chapeau..)","Récompenses financières (coupons gratuits, parking, ticket)","Connaître le groupe sanguin","Avoir un bilan de santé","Appréciation des donateurs via les médias (journaux, site Web, radio, télévision)"]
for collonne in columnsincit :
    incit[collonne]=incit[collonne].apply(remove_arabic).str.strip()
    #com
collcom=["Via une application mobile (qui contient des informations importantes pour donner et recevoir du sang)",
"Des courriers (mail), des SMS )",
"Les réseaux sociaux",
"Les médias (TV, radio…)",
"La presse locale, nationale",
"Présence dans: événements, grandes surfaces, foires, souk hebdomadaire",
"Via d'autre solutions innovantes de gestion de donneurs (les rendez-vous, gestion et traçabilité..)","Question 9 - Quelle est la probabilité que vous fassiez des dons de sang dans le futur ?"]
for champ in collcom :
    com[champ]=com[champ].apply(remove_arabic).str.strip()
    #qst3.4
columnsqst3_4=['Satisfaction_Accueil','Satisfaction_Salle_Attente','Satisfaction_Pendant_Don','Satisfaction_Apres_Don']

for k in columnsqst3_4 :
    qst3_4[k]=qst3_4[k].apply(remove_arabic).str.strip()
    #experience
columnsexp=['J\'ai beaucoup appris et bénéficié','Je pense que je donnerai plus souvent ','C’est très différent des dons que j’ai pu faire à d\'autres associations','C’est devenu une habitude','Je me sens de plus en plus à l’aise lorsque je vais donner mon sang']
for cull in columnsexp :
    exp[cull]=exp[cull].apply(remove_arabic).str.strip()
modification={"Extrêmement important":"Extrêmement Important","Très important":"Très Important","peu important":"Peu Important","Sans opinion":""}
exp=exp.replace(modification)
motiv=motiv.replace(modification)
qst_6=qst_6.replace(modification)
com=com.replace(modification)

#3.sauvegarder les tableaux apres les changement dans un autre fichier excel nommé data.
writer = pd.ExcelWriter(r"C:\Users\lenovo\Desktop\data22.xlsx", engine='xlsxwriter')
profil.to_excel(writer, sheet_name='Profil Participants', startrow=0, index=False)
don.to_excel(writer, sheet_name="Don du sang", startrow=0, index=False)
qst1_1.to_excel(writer, sheet_name="Question 1.1", startrow=0, index=False)
motiv.to_excel(writer, sheet_name="Motivation Qst4_5", startrow=0, index=False)
qst_6.to_excel(writer, sheet_name="Qst 6 idée fausse", startrow=0, index=False)
incit.to_excel(writer, sheet_name="Incitation Sst7", startrow=0, index=False)
com.to_excel(writer, sheet_name="Communication Qst 8_9_10", startrow=0, index=False)
mngmt.to_excel(writer, sheet_name="Management des donneurs", startrow=0, index=False)
mngmt2.to_excel(writer, sheet_name="Management 2", startrow=0, index=False)
qst3_4.to_excel(writer, sheet_name="Qst 3_4 Satisfaction", startrow=0, index=False)
exp.to_excel(writer, sheet_name=" 3.5 expérience leçon  tiré des", startrow=0, index=False)
writer.save()
writer.close()
