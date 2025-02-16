class TextShape {
    constructor(element) {
        this.element = element;
        this.id = element.getObjectId();
        this.isShape = element.getPageElementType() === SlidesApp.PageElementType.SHAPE;
        this.isTextBox = this.isShape ? element.asShape().getShapeType() === SlidesApp.ShapeType.TEXT_BOX : null;
        this.textRange = this.isTextBox ? element.asShape().getText() : null;
        this.hasText = this.textRange && this.textRange.asString().trim() !== "";
        this.top = element.getTop();
        this.left = element.getLeft();
    }

    isUppermost(shapes) {
        var sortedTextShapes = shapes.map(element => new TextShape(element))
            .filter(shape => shape.hasText)
            .sort((a, b) => a.top - b.top);
        return sortedTextShapes[0].id === this.id;
    }

    getTextStyle() {
        return this.hasText ? this.textRange.getTextStyle() : null;
    }

    setTextStyle(style) {
        if (this.hasText) {
            if (style.fontFamily) this.textRange.getTextStyle().setFontFamily(style.fontFamily);
            if (style.foregroundColor) this.textRange.getTextStyle().setForegroundColor(style.foregroundColor);
            if (style.fontSize) this.textRange.getTextStyle().setFontSize(style.fontSize);
        }
    }

    setPosition(top, left) {
        if (this.isShape) {
            this.element.setTop(top);
            this.element.setLeft(left);
        }
    }
}
