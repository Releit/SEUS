UPDATE
  T_Lieferscheine
SET
  T_Lieferscheine.Nr = Int([nr] / 10000) * 100000 + [Nr] Mod 10000;
