import { CAT_WALK_GIF_DATA_URI } from "./cat-walk-gif.ts";

/**
 * Card Service fixed footers only take buttons, so the cat can't truly be pinned;
 * it sits in its own last section, at the bottom of the card content.
 */
export function buildCatSection(): GoogleAppsScript.Card_Service.CardSection {
  return CardService.newCardSection().addWidget(
    CardService.newImage()
      .setImageUrl(CAT_WALK_GIF_DATA_URI)
      .setAltText("A black pixel-art cat walking across the card"),
  );
}
