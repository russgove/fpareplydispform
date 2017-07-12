const adalConfig: adal.Config = {
 
  clientId: 'f8f8d2ad-7c9d-4aac-80eb-3f00a263c879',//this can only read greaph
  tenant: 'common',
  extraQueryParameter: 'nux=1',
  endpoints: {
    'https://rgove3.sharepoint.com': 'f8f8d2ad-7c9d-4aac-80eb-3f00a263c879'
  },
  postLogoutRedirectUri: window.location.origin,
  cacheLocation: 'sessionStorage',
  popUp: false,
  navigateToLoginRequestUrl:false

};

export default adalConfig;
//The app identifier has been successfully created.
//Client Id:  	ac1402af-ce61-417a-8432-1ef8416c5ce5
//Client Secret:  	hbrZgg/SM5fKJW5R8R7D01lJfiN7K15qi47aShfHzSc=
//Title:  	outlookwebapp
//App Domain:  	www.google.com
//Redirect URI:  	https://localhost:3000