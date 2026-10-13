import { getEmailDomain } from "../email-utils.ts";
import { log } from "../log.ts";

export function getMyEmail(): string {
  const cache = CacheService.getUserCache();
  const cached = cache.get("my_email");
  if (cached) return cached;

  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      cache.put("my_email", email, 3600);
    }
    return email;
  } catch (err) {
    log.error("getActiveUser email failed", err);
    return "";
  }
}

export function getMyDomain(): string {
  return getEmailDomain(getMyEmail());
}
