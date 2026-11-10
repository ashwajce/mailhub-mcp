import { Account, Provider } from "../types.js";
import { gmail } from "./gmail.js";
import { outlook } from "./outlook.js";

export function providerFor(account: Account): Provider {
  return account.provider === "gmail" ? gmail : outlook;
}
