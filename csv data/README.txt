Seat-layout CSV datasets generated from the three legacy Excel (.xls) files, with a small manual extension for the College Feast dessert layout.

FILES
- *_seats.csv: one row per O-marked seat. seat_id is the Excel cell address.
- *_tables.csv: one row per table section, with its ordered seat list.
- *_seat_pairs.csv: ordered/symmetric close-seat pairs, directly usable as C and w_sr.
- all_layouts_*.csv: combined datasets for all layouts.
- layout_summary.csv: validation counts.

PAIR RULES
- same-side or cycle neighbours: weight 1.0
- direct opposite and the two seats immediately beside the opposite seat: weight 0.5
- explicit corner/junction neighbours in the E-shaped Upper Hall layout: weight 1.0

IMPORTANT
- Hall 110 contains 112 cells marked O in the source Excel file. All 112 are preserved.
- Master and VIP seat columns are intentionally blank because those roles are not identified by the layout drawings.
- Hall high-table end seats are marked END; other high-table seats are marked MAIN_SIDE. Centre/VIP classifications should be confirmed separately.

MANUAL COLLEGE FEAST EXTENSION
- The original Upper Hall layout contains 45 O-marked seats.
- Following the supervisor's guidance, five seats were added manually for the College Feast dessert arrangement: H16, H17, H24, H25 and I31.
- These seats are assigned to the RIGHT_VERTICAL section.
- The extended Upper Hall dataset contains 50 seats; together with the 60-seat Combination Room, the dessert capacity is 110.
- detected_O_cells remains 45 because the five added seats are not O-marked cells in the original Excel file.
- The original H31--J30 direct corner pair was replaced by H31--I31 and I31--J30.
