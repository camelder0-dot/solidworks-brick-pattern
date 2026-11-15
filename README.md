# Brick Pattern for SOLIDWORKS

A small SOLIDWORKS macro that **engraves a brick pattern into a flat surface**,
like the top of a wooden board, so it can be CNC-routed or laser-engraved.

## What it does

Select a flat face, run the macro, and answer five quick prompts:

| Setting | Default | Meaning |
|---|---|---|
| Brick length | 40 mm | Length of each brick |
| Brick height | 15 mm | Height of each row |
| Mortar line width | 2 mm | Width of the lines between bricks |
| Carve depth | 1 mm | How deep the pattern is cut into the surface |
| Carve style | 1 | **1** = carve the mortar lines (bricks stay raised) · **2** = carve the bricks (mortar stays raised) |

Rows are staggered (running bond), like a real brick wall. The pattern is
clipped to the face automatically.

The result is ordinary **Cut-Extrude** features in the FeatureManager tree
(`Brick Mortar Rows` + `Brick Mortar Joints`, or `Brick Pockets`).
Double-click one to change the depth afterward, or delete it to start over.

## Install (one time)

1. In SOLIDWORKS, go to **Tools → Macro → New…** and save it as `BrickPattern.swp`.
2. The VBA editor opens. Choose **File → Import File…** and pick `BrickPattern.bas` from this repo.
3. In the Project panel on the left, delete the empty module SOLIDWORKS created (usually `Module1`), so only `BrickPattern` is left.
4. Save (**Ctrl+S**) and close the VBA editor.

**Optional toolbar button:** go to **Tools → Customize → Commands → Macro**, drag the
**New Macro Button** onto a toolbar, and point it at `BrickPattern.swp`.

## Use

1. Open your part (for example, a board made with a Boss-Extrude).
2. Click the face you want to engrave.
3. Go to **Tools → Macro → Run…**, pick `BrickPattern.swp`, and fill in the prompts (or click your toolbar button).

## Changing the defaults

The default values are at the top of `BrickPattern.bas`:

```vb
Private Const DEF_BRICK_LENGTH As Double = 40   ' mm
Private Const DEF_BRICK_HEIGHT As Double = 15   ' mm
Private Const DEF_MORTAR_WIDTH As Double = 2    ' mm
Private Const DEF_CARVE_DEPTH As Double = 1     ' mm
Private Const DEF_CARVE_STYLE As Long = 1       ' 1 = carve mortar, 2 = carve bricks
```

## Tips

- Bricks follow the face's sketch axes. If they run the wrong way, swap the brick length and height values.
- Keep the carve depth smaller than the board thickness.
- Very small bricks on a big face make many rectangles and can be slow.
