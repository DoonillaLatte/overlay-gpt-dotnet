# PowerPoint 스타일을 HTML로 변환하기

## 1. 텍스트 스타일 변환

### 기본 텍스트 스타일
```css
- 글꼴 크기: font-size: {size}pt
- 글꼴 이름: font-family: {fontName}
- 글자 굵기: font-weight: Bold/Normal
- 이탤릭: <i>태그
- 밑줄: <u>태그
- 취소선: <s>태그
- 텍스트 색상: color: #{rgbColor}
```

### 배경 스타일
```css
- 배경색: background-color: rgba(r, g, b, alpha)
- 하이라이트: background-color: rgb(r, g, b)
```

## 2. 정렬 스타일

### 수평 정렬
```css
- center: justify-content: center
- right: justify-content: flex-end
- left: justify-content: flex-start
```

### 수직 정렬
```css
- middle: align-items: center
- bottom: align-items: flex-end
- top: align-items: flex-start
```

## 3. 도형 스타일

### 기본 도형 속성
```css
- 위치: position: absolute
- 좌표: left: {x}px, top: {y}px
- 크기: width: {width}px, height: {height}px
- 회전: transform: rotate({angle}deg)
```

### 테두리와 효과
```css
- 테두리: border: {weight}px {style} {color}
- 그림자: box-shadow: {x}px {y}px {blur}px rgba(r,g,b,alpha)
- 모서리 둥글기: border-radius: {radius}px
- Z-인덱스: z-index: {position}
```

## 4. HTML 태그 변환

### 도형 타입별 태그
```html
- 자동 도형: <div>
- 그림: <img>
- 텍스트 상자: <div>
- 선: <div>
- 차트: <div>
- 표: <table>
- SmartArt: <div>
```

## 5. 특수 효과

### 3D 효과
```css
- transform-style: preserve-3d
- perspective: 1000px
- transform: rotateX() rotateY()
```

### 그라데이션
```css
- background: linear-gradient(direction, color-stops)
```

## 6. 변환 처리 메서드

주요 변환 메서드:
- `ConvertShapeToHtml()`: 도형을 HTML로 변환
- `GetStyledText()`: 텍스트 스타일 적용
- `GetTextStyleString()`: 텍스트 스타일 문자열 생성
- `GetShapeStyleString()`: 도형 스타일 문자열 생성 