import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import notificationTemplate from "./adaptiveCards/notification-default.json";
import { CardData } from "./cardModels";
import { cosmosStore, bot } from "./internal/initialize";

// An Azure Function HTTP trigger.
//
// This endpoint is provided by your application to listen to events. You can configure
// your IT processes, other applications, background tasks, etc - to POST events to this
// endpoint.
//
// In response to events, this function sends Adaptive Cards to Teams. You can update the logic in this function
// to suit your needs. You can enrich the event with additional data and send an Adaptive Card as required.
//
// You can add authentication / authorization for this API. Refer to
// https://aka.ms/teamsfx-notification for more details.
const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  // By default this function will iterate all the installation points and send an Adaptive Card
  // to every installation.
  const pageSize = 10;
  let continuationToken: string | undefined;
  do {
    const pagedInstallations = await bot.notification.getPagedInstallations(pageSize, continuationToken);
    continuationToken = pagedInstallations.continuationToken;
    const targets = pagedInstallations.data;
    for (const target of targets) {
      await target.sendAdaptiveCard(
        AdaptiveCards.declare<CardData>(notificationTemplate).render({
          title: "New Event Occurred!",
          appName: "Contoso App Notification",
          description: `This is a sample http-triggered notification to ${target.type}`,
          notificationUrl: "https://www.adaptivecards.io/",
        })
      );
    }
  } while (continuationToken);


  // You can also find someone and notify the individual person
  const member = await bot.notification.findMember(
    async (m) => m.account.email === "someone@contoso.com"
  );
  await member?.sendAdaptiveCard(
    AdaptiveCards.declare<CardData>(notificationTemplate).render({
      title: "New Event Occurred!",
      appName: "Contoso App Notification",
      description: `This is a sample http-triggered notification to ${member.type}`,
      notificationUrl: "https://www.adaptivecards.io/",
    })
  );

  // You can also find someone and notify the individual person
  const conversationReference = await cosmosStore.getUserByUserId("92d17b4c-0325-46cb-9b5f-c3ea6964bef6");
  // const conversationReference = await cosmosStore.getUserByUserEmail("someone@consto.com");
  
  const installation = bot.notification.buildTeamsBotInstallation(conversationReference);
  await installation?.sendAdaptiveCard(
    AdaptiveCards.declare<CardData>(notificationTemplate).render({
      title: "New Event Only For You!",
      appName: "Contoso App Notification",
      description: `This is a sample http-triggered notification to ${installation.type}`,
      notificationUrl: "https://www.adaptivecards.io/",
    })
  );

  context.res = {};
};

export default httpTrigger;
