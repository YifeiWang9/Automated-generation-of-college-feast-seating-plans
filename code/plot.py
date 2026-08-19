from pathlib import Path
import csv

import matplotlib

# Allows the script to run without opening a graphical window.
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle, Circle

# File paths
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR / "csv data"

FIG_DIR = SCRIPT_DIR / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

SEAT_FILE = DATA_DIR / "all_layouts_seats.csv"


def read_csv(path: Path) -> list[dict[str, str]]:
    """Read a UTF-8 CSV file and return a list of dictionaries."""
    if not path.exists():
        raise FileNotFoundError(
            f"Could not find {path}. Check the DATA_DIR setting."
        )

    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


seats = read_csv(SEAT_FILE)



# Academic monochrome formatting
MARKERS = ["o", "s", "^", "v", "D", "P", "X"]

GREYS = [
    "0.15",
    "0.28",
    "0.40",
    "0.52",
    "0.64",
    "0.74",
    "0.84",
]

SECTION_ORDER = {
    "hall_110": [
        "HIGH_TABLE",
        "LEFT_UPPER",
        "MIDDLE_UPPER",
        "RIGHT_UPPER",
        "LEFT_LOWER",
        "MIDDLE_LOWER",
        "RIGHT_LOWER",
    ],
    "combination_60": [
        "TOP_TABLE",
        "ROUND_TABLE",
        "OCTAGON_MIDDLE",
        "RIGHT_TABLE",
        "OCTAGON_LOWER",
    ],
    "upper_45": [
        "TOP_BAR",
        "MIDDLE_BAR",
        "BOTTOM_BAR",
        "RIGHT_VERTICAL",
    ],
}
 

def pretty(text: str) -> str:
    """Convert labels such as LEFT_UPPER into Left Upper."""
    return text.replace("_", " ").title()




def draw_layout(
    layout: str,
    output_name: str,
    title: str,
    added_seats: set[str] | None = None,
    show_seat_labels: bool = False,
) -> None:
    """
    Draw one physical layout.

    Sections are distinguished using marker shape and grey level rather than
    bright colours, so the figure remains readable when printed in greyscale.
    """
    added_seats = set(added_seats or set())

    layout_rows = [
        row for row in seats
        if row["layout"] == layout
    ]

    if not layout_rows:
        raise ValueError(f"No seats were found for layout {layout!r}.")

    section_order = [
        section
        for section in SECTION_ORDER[layout]
        if any(
            row["section_id"] == section
            for row in layout_rows
        )
    ]

    figure, ax = plt.subplots(figsize=(8.4, 7.2))

    legend_handles: list[Line2D] = []

    for index, section in enumerate(section_order):
        section_rows = [
            row for row in layout_rows
            if row["section_id"] == section
        ]

        ordinary_rows = [
            row for row in section_rows
            if row["seat_id"] not in added_seats
        ]

        added_rows = [
            row for row in section_rows
            if row["seat_id"] in added_seats
        ]

        marker = MARKERS[index % len(MARKERS)]
        grey = GREYS[index % len(GREYS)]

        if ordinary_rows:
            ax.scatter(
                [int(row["excel_col"]) for row in ordinary_rows],
                [int(row["excel_row"]) for row in ordinary_rows],
                s=52,
                marker=marker,
                facecolor=grey,
                edgecolor="black",
                linewidth=0.55,
                zorder=3,
            )

        if added_rows:
            ax.scatter(
                [int(row["excel_col"]) for row in added_rows],
                [int(row["excel_row"]) for row in added_rows],
                s=115,
                marker="*",
                facecolor="white",
                edgecolor="black",
                linewidth=1.1,
                zorder=5,
            )

 

        legend_handles.append(
            Line2D(
                [0],
                [0],
                marker=marker,
                linestyle="None",
                markerfacecolor=grey,
                markeredgecolor="black",
                markersize=7,
                label=pretty(section),
            )
        )
    
    if show_seat_labels:
        for row in layout_rows:
            seat_id = row["seat_id"]
            x_value = int(row["excel_col"])
            y_value = int(row["excel_row"])

            if seat_id in added_seats:
                # Put manually added-seat labels to the right of the star.
                x_offset = 5
                y_offset = -5
                horizontal_alignment = "left"
                vertical_alignment = "center"
                font_size = 6.5
            else:
                # Put ordinary seat labels slightly above the marker.
                x_offset = 0
                y_offset = 5
                horizontal_alignment = "center"
                vertical_alignment = "bottom"
                font_size = 5.2

            ax.annotate(
                seat_id,
                (x_value, y_value),
                xytext=(x_offset, y_offset),
                textcoords="offset points",
                ha=horizontal_alignment,
                va=vertical_alignment,
                fontsize=font_size,
                zorder=6,
            )

    if added_seats:
        legend_handles.append(
            Line2D(
                [0],
                [0],
                marker="*",
                linestyle="None",
                markerfacecolor="white",
                markeredgecolor="black",
                markersize=10,
                label="Manually added seat",
            )
        )

    all_x = [
        int(row["excel_col"])
        for row in layout_rows
    ]

    all_y = [
        int(row["excel_row"])
        for row in layout_rows
    ]

    ax.set_xlim(
        min(all_x) - 1.5,
        max(all_x) + 1.5,
    )

    # Reverse the vertical direction to match Excel row positions.
    ax.set_ylim(
        max(all_y) + 1.8,
        min(all_y) - 1.8,
    )

    ax.set_aspect("equal", adjustable="box")
    ax.set_xlabel("Excel column coordinate")
    ax.set_ylabel("Excel row coordinate")
    ax.set_title(title, fontsize=13, pad=12)

    ax.grid(
        color="0.92",
        linewidth=0.6,
    )

    ax.legend(
        handles=legend_handles,
        loc="upper left",
        bbox_to_anchor=(1.02, 1.0),
        frameon=False,
        fontsize=9,
    )

    for spine in ax.spines.values():
        spine.set_linewidth(0.8)

    output_path = FIG_DIR / output_name

    figure.savefig(
        output_path,
        dpi=300,
        bbox_inches="tight",
    )

    plt.close(figure)



# High Table figure
def draw_high_table() -> None:
    """Draw the implemented Master, VIP and High Table end classifications."""
    high_table_rows = [
        row
        for row in seats
        if (
            row["layout"] == "hall_110"
            and row["is_high_table"].strip().lower()
            in {"1", "true"}
        )
    ]

    master_seat = "I8"

    vip_seats = {
        "H8",
        "K8",
        "H12",
        "I12",
        "J12",
        "K12",
    }

    end_seats = {
        row["seat_id"]
        for row in high_table_rows
        if row["high_table_zone"] == "END"
    }

    figure, ax = plt.subplots(figsize=(9.0, 3.8))

    for row in high_table_rows:
        seat_id = row["seat_id"]
        x_value = int(row["excel_col"])
        y_value = int(row["excel_row"])

        if seat_id == master_seat:
            marker = "s"
            face_colour = "0.10"
            marker_size = 95

        elif seat_id in vip_seats:
            marker = "o"
            face_colour = "0.55"
            marker_size = 78

        elif seat_id in end_seats:
            marker = "^"
            face_colour = "white"
            marker_size = 85

        else:
            marker = "o"
            face_colour = "white"
            marker_size = 58

        ax.scatter(
            [x_value],
            [y_value],
            s=marker_size,
            marker=marker,
            facecolor=face_colour,
            edgecolor="black",
            linewidth=0.8,
            zorder=3,
        )

        ax.annotate(
            seat_id,
            (x_value, y_value),
            xytext=(0, 7),
            textcoords="offset points",
            ha="center",
            fontsize=10,
        )

    legend_handles = [
        Line2D(
            [0],
            [0],
            marker="s",
            linestyle="None",
            markerfacecolor="0.10",
            markeredgecolor="black",
            markersize=8,
            label="Master (I8)",
        ),
        Line2D(
            [0],
            [0],
            marker="o",
            linestyle="None",
            markerfacecolor="0.55",
            markeredgecolor="black",
            markersize=7,
            label="VIP seat",
        ),
        Line2D(
            [0],
            [0],
            marker="^",
            linestyle="None",
            markerfacecolor="white",
            markeredgecolor="black",
            markersize=8,
            label="High Table end",
        ),
        Line2D(
            [0],
            [0],
            marker="o",
            linestyle="None",
            markerfacecolor="white",
            markeredgecolor="black",
            markersize=7,
            label="Other High Table seat",
        ),
    ]

    ax.set_aspect("equal", adjustable="box")
    ax.invert_yaxis()

    ax.set_title(
        "High Table classifications used in the model",
        fontsize=14,
        pad=10,
    )

    ax.set_xlabel("Excel column coordinate")
    ax.set_ylabel("Excel row coordinate")

    ax.grid(
        color="0.92",
        linewidth=0.6,
    )

    ax.legend(
        handles=legend_handles,
        loc="upper left",
        bbox_to_anchor=(1.02, 1.0),
        frameon=False,
        fontsize=10,
    )

    figure.savefig(
        FIG_DIR / "high_table_classification.png",
        dpi=300,
        bbox_inches="tight",
    )

    plt.close(figure)



# Close-seat weight schematic
def draw_close_seat_weights() -> None:
    """Draw a clean schematic of the weights 1 and 0.5."""
    figure, ax = plt.subplots(figsize=(9.2, 5.0))

    ax.set_aspect("equal")
    ax.axis("off")

    top_row = [
        (1.0, 3.0),
        (2.5, 3.0),
        (4.0, 3.0),
        (5.5, 3.0),
        (7.0, 3.0),
    ]

    bottom_row = [
        (1.0, 1.0),
        (2.5, 1.0),
        (4.0, 1.0),
        (5.5, 1.0),
        (7.0, 1.0),
    ]

    # Draw all seats.
    for x_value, y_value in top_row + bottom_row:
        ax.add_patch(
            Circle(
                (x_value, y_value),
                0.17,
                facecolor="white",
                edgecolor="black",
                linewidth=1.0,
            )
        )

    # Focal seat.
    ax.add_patch(
        Circle(
            (4.0, 3.0),
            0.20,
            facecolor="0.20",
            edgecolor="black",
            linewidth=1.1,
        )
    )

    ax.text(
        4.0,
        3.48,
        "focal seat",
        ha="center",
        fontsize=11,
    )

    # Same-side direct neighbours: weight 1.
    for neighbour_x in (2.5, 5.5):
        ax.plot(
            [4.0, neighbour_x],
            [3.0, 3.0],
            color="black",
            linewidth=1.8,
        )

        ax.text(
            (4.0 + neighbour_x) / 2,
            3.22,
            "1",
            ha="center",
            fontsize=10.5,
        )

    # Opposite and near-opposite seats: weight 0.5.
    for opposite_x in (2.5, 4.0, 5.5):
        ax.plot(
            [4.0, opposite_x],
            [3.0, 1.0],
            color="0.35",
            linewidth=1.3,
            linestyle="--",
        )

        ax.text(
            (4.0 + opposite_x) / 2 + 0.06,
            1.96,
            "0.5",
            fontsize=10,
        )

    # Light table edges.
    ax.plot(
        [0.55, 7.45],
        [2.55, 2.55],
        color="0.80",
        linewidth=2.0,
    )

    ax.plot(
        [0.55, 7.45],
        [1.45, 1.45],
        color="0.80",
        linewidth=2.0,
    )

    ax.text(
        4.0,
        0.26,
        (
            "Pairs with no edge are not evaluated by "
            "the close-seat objective."
        ),
        ha="center",
        fontsize=10.5,
    )

    ax.set_xlim(0.3, 7.7)
    ax.set_ylim(0.0, 3.8)

    ax.set_title(
        "Modelling weights used for close-seat relations",
        fontsize=14,
        pad=10,
    )

    figure.savefig(
        FIG_DIR / "close_seat_weights.png",
        dpi=300,
        bbox_inches="tight",
    )

    plt.close(figure)



# Near and with schematic
def draw_near_with_definition() -> None:
    fig, ax = plt.subplots(figsize=(8.6, 6.0))

    # draw table body
    table = Rectangle((2.2, 1.6), 5.6, 2.8,
                      facecolor="#f2f2f2", edgecolor="black", linewidth=1.5)
    ax.add_patch(table)

    # focal seat and surrounding seats
    # coordinates are chosen to look like a local fragment of a long table
    seats = {
        "s1": (2.8, 1.2),
        "s2": (4.0, 1.2),
        "focal": (5.0, 1.2),
        "s3": (6.0, 1.2),
        "s4": (7.2, 1.2),

        "s5": (2.8, 4.8),
        "s6": (4.0, 4.8),
        "s7": (5.0, 4.8),
        "s8": (6.0, 4.8),
        "s9": (7.2, 4.8),

        "s10": (2.0, 2.2),
        "s11": (2.0, 3.0),
        "s12": (2.0, 3.8),

        "s13": (8.0, 2.2),
        "s14": (8.0, 3.0),
        "s15": (8.0, 3.8),
    }

    # Example sets:
    # with = seats most closely associated with focal seat
    with_set = {"s2", "s3", "s6", "s7", "s8"}
    # near includes with plus a wider neighbourhood
    near_only_set = {"s1", "s4", "s9", "s14"}

    # draw seats
    r = 0.16
    for name, (x, y) in seats.items():
        if name == "focal":
            face = "black"
            edge = "black"
            lw = 1.5
        elif name in with_set:
            face = "#7f7f7f"     
            edge = "black"
            lw = 1.2
        elif name in near_only_set:
            face = "white"
            edge = "black"
            lw = 1.2
        else:
            face = "white"
            edge = "#999999"
            lw = 1.0

        seat = Circle((x, y), r, facecolor=face, edgecolor=edge, linewidth=lw)
        ax.add_patch(seat)

    # label focal seat
    ax.text(4.5, 0.78, "focal seat", fontsize=12, va="center")

    # labels for with / near
    ax.text(6.55, 2.35, '"with"', fontsize=12)
    ax.text(8.55, 3.0, '"near" only', fontsize=12)

    # optional guide boxes
    # box around with seats
    with_box = Rectangle((3.55, 0.7), 2.8, 4.6,
                         fill=False, linestyle="solid",
                         linewidth=1.3, edgecolor="black")
    ax.add_patch(with_box)

    # wider dashed box around near seats
    near_box = Rectangle((1.55, 0.7), 6.9, 4.6,
                         fill=False, linestyle="dashed",
                         linewidth=1.3, edgecolor="#555555")
    ax.add_patch(near_box)

    ax.text(4.1, 5.38, '"with" neighbourhood', fontsize=12)
    ax.text(1.65, 5.78, '"near" neighbourhood', fontsize=12)

    # explanation
    ax.text(5.0, 0.12,
            'This figure is illustrative only. The model uses explicit seat-pair sets,\n'
            r'with $C_k^{\mathrm{with}} \subseteq C_k^{\mathrm{near}}$.',
            ha="center", va="bottom", fontsize=11)

    ax.set_xlim(0.8, 9.2)
    ax.set_ylim(0.0, 6.2)
    ax.set_aspect("equal")
    ax.axis("off")

    plt.tight_layout()
    fig.savefig(
        FIG_DIR / "near_with_definition.png",
        dpi=300,
        bbox_inches="tight",
    )
    
    plt.close(fig)



# Generate all figures
draw_layout(
    layout="hall_110",
    output_name="hall_sections.png",
    title="Hall layout and section partition",
)

draw_layout(
    layout="combination_60",
    output_name="combination_room_sections.png",
    title="Combination Room dessert layout",
)

draw_layout(
    layout="upper_45",
    output_name="upper_hall_extension.png",
    title="Extended Upper Hall dessert layout",
    added_seats={
        "H16",
        "H17",
        "H24",
        "H25",
        "I31",
    },
    show_seat_labels=True,
)

draw_high_table()
draw_close_seat_weights()
draw_near_with_definition()

print(f"Figures written to: {FIG_DIR.resolve()}")