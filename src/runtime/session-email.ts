import { getEmailDomain } from "../email-utils.ts";

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
    console.log(
      `[sequel] getActiveUser email failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
    return "";
  }
}

export function getMyDomain(): string {
  return getEmailDomain(getMyEmail());
}
