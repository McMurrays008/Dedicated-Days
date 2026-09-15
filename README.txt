DEDICATED DAYS — MANUAL MORNING / AUTOMATED STATUS UPDATE

Replace these THREE files in the McMurrays008/Dedicated-Days GitHub repo:
  collector.py
  server.py
  requirements.txt

What changes:
- No 08:00 full/Integration refresh.
- No Integration export is required.
- At about 09:30, open /admin and upload the morning TPN Dedicated Day Check .xlsx.
- The morning file fixes the day's delivery population using the 10 September method.
- Delivery Depot 8 is INCLUDED.
- At 10:00, 12:00, 14:00, 16:00 and 18:00 Europe/London, Browse runs automatically.
- Browse updates Status only for Dockets already in the morning population.
- Status refreshes cannot add or remove deliveries.
- The existing Update Status Now button remains available for an extra manual refresh.

After GitHub deployment:
1. Open https://dedicated-days.onrender.com/admin
2. Paste the existing REFRESH_TOKEN.
3. Choose the morning TPN Dedicated Day Check .xlsx.
4. Click Import Morning Check.
