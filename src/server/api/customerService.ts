import express = require("express");
import passport = require("passport");
import { BearerStrategy, IBearerStrategyOption, ITokenPayload, VerifyCallback } from "passport-azure-ad";
import { ICustomer } from "../../model/ICustomer";
import { retrieveConfig } from "./appConfigService";
import { getCustomer } from "./azTableService";

export const customerService  = (options: any) =>  {
    const router = express.Router();

    // Set up the Bearer Strategy
  const bearerStrategy = new BearerStrategy({
    identityMetadata: "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
    clientID: process.env.MICROSOFT_APP_ID as string,
    audience: `api://${process.env.PUBLIC_HOSTNAME}/${process.env.MICROSOFT_APP_ID}`,
    loggingLevel: "warn",
    validateIssuer: false,
    passReqToCallback: false
  } as IBearerStrategyOption,
    (token: ITokenPayload, done: VerifyCallback) => {
        done(null, { tid: token.tid, name: token.name, upn: token.upn }, token);
    }
  );
  const pass = new passport.Passport();
  router.use(pass.initialize());
  pass.use(bearerStrategy);

  router.get(
      "/customer/:meetingID",
      async (req: any, res: express.Response, next: express.NextFunction) => {
          const meetingID: any = req.params.meetingID;
          // const config: ICustomer = await retrieveConfig(meetingID);
          const config: ICustomer = await getCustomer(meetingID);
          res.json(config);
    });
    
    return router;
};