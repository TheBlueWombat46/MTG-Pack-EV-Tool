# MTG Pack EV Tool

An expected value (EV) model that answers a store owner's question:
**is a Magic: The Gathering booster pack worth more sealed, or opened and sold as singles?**

Built in Google Sheets with Apps Script. Current card prices come from the
[Scryfall API](https://scryfall.com/docs/api), and sheet formulas combine them
with booster pull rates to estimate value per pack.

## How it works
1. **Price refresh (Apps Script):** pulls every paper printing in a chosen set,
   with separate rows for normal, foil, and etched finishes.
2. **Booster configuration:** each pack slot has a card count and pull
   probabilities.
3. **Two kinds of EV:**
   - **Theoretical EV:** the market value of the expected cards in a pack.
   - **Realizable EV:** what you'd actually take home after marketplace fees,
     selling below market price, per-card shipping costs, and a minimum
     listing value (cards under $1 aren't worth listing).
4. **Dashboard:** compares both EVs with pack cost to show expected profit and ROI.

## Current status
The model is built and validated on **Core Set 2021 (M21) draft boosters**.
At a $6 pack cost, theoretical EV is almost even with cost, but realizable EV is far
below it. Most of the value is in bulk commons and uncommons that can't be
sold profitably, so M21 draft boosters are worth more sealed.

## Validation and fixes
- Found that realizable prices were reading the wrong column, which zeroed out
  every non-land card. Traced the error and fixed it in the script.
- Checked card pools against official product contents and found
  Collector-Booster-only extended-art cards inflating draft booster EV.
  Excluded them from draft pools.

## In progress: pool-based model
The current model groups cards by rarity. I'm rebuilding it around **card pools**
so each product (Draft, Collector, Play Booster) can define exactly which cards
each slot draws from. This handles sets where cards appear in some products but
not others. Work happens on the `pool-model` branch.

## Tech
Google Sheets · Google Apps Script (JavaScript) · Scryfall REST API
