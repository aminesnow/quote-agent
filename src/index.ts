import { App } from '@microsoft/teams.apps';
import { BotBuilderPlugin } from '@microsoft/teams.botbuilder';
import { DevtoolsPlugin } from '@microsoft/teams.dev';
import { ConfigurationBotFrameworkAuthentication, ConfigurationServiceClientCredentialFactory, TeamsActivityHandler, TurnContext } from 'botbuilder';
import { CloudAdapter } from 'botbuilder';

class ActivityHandler extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (ctx, next) => {
      console.log('BotBuilder handler received message:', ctx.activity.text);
      await ctx.sendActivity('hi from botbuilder...');
      await next();
    });
  }

  async run(context: TurnContext): Promise<void> {
    console.log('BotBuilder handler received message:', context.activity.text);
    await context.sendActivity('hi from botbuilder...');
  }
}

const handler = new ActivityHandler();

const adapter = new CloudAdapter(
  new ConfigurationBotFrameworkAuthentication(
    {},
    new ConfigurationServiceClientCredentialFactory({
      MicrosoftAppType: 'MultiTenant',
      MicrosoftAppTenantId: "d0afe711-d9a2-46fd-9d6c-fd8b481b5a37",
    })
  )
);

// Create BotBuilder plugin with simple ActivityHandler
const botBuilderPlugin = new BotBuilderPlugin({ adapter, handler });

const app = new App({
  plugins: [
    new DevtoolsPlugin(),
    botBuilderPlugin,
  ],
});

app.on('message', async ({ send, activity }) => {
  await send({ type: 'typing' });
  await send(`you said: "${activity.text}"`);
});

(async () => {
  await app.start(+(process.env.PORT || 3978));
})();
