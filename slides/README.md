# YAML Slide Format

The deck generator reads `slide*.yaml` files in this directory and renders them
in numeric order.

Run:

```bash
python3 generate_observability_stack_deck.py
```

Or with custom paths:

```bash
python3 generate_observability_stack_deck.py --slides-dir slides --output custom-deck.pptx
```

## Slide structure

Each file looks like:

```yaml
background: white
elements:
  - kind: title
    text: My Slide Title

  - kind: subtitle
    text: Optional subtitle

  - kind: bullets
    x: 0.8
    y: 1.6
    w: 6.0
    h: 4.0
    font_size: 20
    bullets:
      - "Quote strings that contain a colon: like this."

  - kind: box
    x: 7.0
    y: 1.6
    w: 4.0
    h: 1.0
    text: Callout box
    fill: light_fill
    font_size: 18
    bold: true
```

Coordinates and sizes are in inches.

## Supported element kinds

- `title`
- `subtitle`
- `textbox`
- `bullets`
- `note`
- `box`
- `arrow`
- `banner`
- `table`

## Color names

You can use these built-in colors:

- `title`
- `accent`
- `secondary`
- `light_fill`
- `pale_green`
- `pale_orange`
- `white`
- `black`

You can also use `#RRGGBB` hex values.
