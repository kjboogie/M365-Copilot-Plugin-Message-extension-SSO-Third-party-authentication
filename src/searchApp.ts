import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
} from "botbuilder";
import * as ACData from "adaptivecards-templating";
import helloWorldCard from "./adaptiveCards/helloWorldCard.json";

export class SearchApp extends TeamsActivityHandler {
  private userState;
  private Token;
  private Expiry;
 
  constructor(userState) {
    super();
    this.userState = userState;
    this.Token = this.userState.createProperty("Token");
    this.Expiry = this.userState.createProperty("Expiry");
  }
  
  staticHtmlPage(): MessagingExtensionResponse {
    return {
      composeExtension:{
        type: "auth",
        suggestedActions: {
          actions: [
            {
              type: "openUrl",
              value:  `https://${process.env.BOT_DOMAIN}/auth.html?clientId=${process.env.CLIENT_ID}`,
              title: "Sign in"
            }
          ],
        },
      }
    };
  }

  async run(context: TurnContext): Promise<void> {
    await super.run(context);
    await this.userState.saveChanges(context, false);
  }

  // this function is to generate tghe token from the auth code. if your third party given direct token than do not need to call these function.
  public async authenticate(context:TurnContext){
    const bearerRegex = /^ey[A-Za-z0-9-_=]+\.[A-Za-z0-9-_=]+\.?[A-Za-z0-9-_.+/=]*$/; // checking if bearer token is valid or not
    let Token = await this.Token.get(context);
    let token = context.activity.value.state
    let isToken = bearerRegex.test(token);
    const authorizationCode = context.activity.value.state;
    console.log("Authorization code:", authorizationCode);
   
   
    // Define the token endpoint and the data for the POST request
    const tokenEndpoint = 'Pass the token generation by auth code url '; // pass the url here
    const postData = {
    code: authorizationCode,
    client_id: process.env.CLIENT_ID,
    client_secret:  process.env.CLIENT_SECRET,
    redirect_uri: `https://${process.env.BOT_DOMAIN}/auth-end.html` ,
   
    };
     // Make the POST request to the token endpoint by passing authorization code
     try {
      const response = await axios.post(tokenEndpoint, postData);
 
      // Log the response data
      console.log('Token response:', response.data.access_token);
      this.Token.set(context,response.data.access_token);
      this.Expiry.set(context,response.data.expires_in);
      
      return ;
      
    } catch (error) {
      console.error('Error fetching token:', error);
    }        ``
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;
    let Token = await this.Token.get(context);
    // let adapter: any = context.adapter;
    // let userClientAction = context.turnState.get(adapter.UserTokenClientKey);
    // const test = await storage.read(["token"])

    if(!Token && !context.activity.value.state){ // here we are checking if we have token and auth code already, if no we will return the 'auth' type compose extension
      const test = this.staticHtmlPage();
      return this.staticHtmlPage();
    }
    else if(context.activity.value.state){ // if we only have auth code then will generate token from it
      await this.authenticate(context);
    }
    // Check if the query is for a search
    if (!query.parameters || !query.parameters[0]) {
          return this.staticHtmlPage();
    }
    // if we have the token then it can do further execution
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const template = new ACData.Template(helloWorldCard);
      const card = template.expand({
        $root: {
          name: obj.package.name,
          description: obj.package.description,
        },
      });
      const preview = CardFactory.heroCard(obj.package.name);
      const attachment = { ...CardFactory.adaptiveCard(card), preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }
}
