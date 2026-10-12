function getUserInfo() {
  try {
    const me = People.People.get('people/me', {
      personFields: 'names,photos,emailAddresses'
    });

    const name = me.names ? me.names[0].displayName : "Utilisateur";
    const photo = me.photos ? me.photos[0].url : "";
    const email = me.emailAddresses ? me.emailAddresses[0].value : Session.getActiveUser().getEmail();


    const result = {
      name: name,
      photoUrl: photo,
      email: email,
      loggedIn: true
    };
    // console.log(result)
    return result;

  } catch (err) {
    Logger.log("Erreur People API : %s", err.message);
    return {
      name: "Utilisateur",
      photoUrl: "",
      email: Session.getActiveUser().getEmail(),
      loggedIn: true
    };
  }
}