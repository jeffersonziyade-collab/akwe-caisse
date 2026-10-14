# Akwè Caisse — backend

Node 18+ · Express · SQLite · JWT.

    npm install
    JWT_SECRET="une-longue-phrase-secrete" npm start     # http://localhost:3000

## Variables d'environnement
- `JWT_SECRET` (obligatoire, 16 caractères min.)
- `PORT`, `DB_PATH` (chemin du fichier SQLite, à mettre sur un disque persistant)
- `BASE_URL` — URL publique (ex. https://akwe.exemple.com), utilisée dans le lien de réinitialisation
- `SMTP_URL` — ex. `smtps://utilisateur:motdepasse@smtp.exemple.com` ; sans elle, le lien de réinitialisation s'affiche dans la console du serveur
- `MAIL_FROM` — expéditeur, ex. `Akwè Caisse <no-reply@exemple.com>`

## API (en-tête `Authorization: Bearer <token>`)
- POST /api/register {name, email, password, consent:true}, /api/login → {token}
- POST /api/forgot {email}, /api/reset {token, password}
- GET /api/data · POST /api/sync · POST /api/sms/parse {text}
- GET /api/report?per=jour|semaine|mois · GET /api/export.xlsx|.pdf?per=…
- DELETE /api/account {password}

## Avant la mise en ligne
- Compléter les champs [entre crochets] de `public/conditions.html` et `public/confidentialite.html`, puis les faire valider par un juriste.
- Servir en HTTPS, sauvegarder `DB_PATH`, tester le parseur SMS avec de vrais messages.
