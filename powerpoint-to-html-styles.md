# PowerPoint 스타일을 HTML로 변환하기

## 1. 텍스트 스타일 변환

### 기본 텍스트 스타일
```css
- 글꼴 크기: font-size: {size}pt
- 글꼴 이름: font-family: {fontName}
- 글자 굵기: font-weight: Bold/Normal
- 이탤릭: font-style: italic
- 밑줄: text-decoration: underline
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
- 그림: <img src='{절대경로}/images/{GUID}.png' alt='Image' />
- 텍스트 상자: <div>
- 선: <div>
- 차트: <div>
- 표: <table>
- SmartArt: <div>
```

### 이미지 처리
```css
- 저장 형식: PNG
- 저장 위치: {프로그램경로}/images/
- 파일명: {GUID}.png
- 참조 방식: 절대 경로 사용
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


## 7. 슬라이드 HTML 변환 구조

### 단일 슬라이드 HTML 구조(주의: <div class='Slide1'></div> 같이 전체를 감싸는 div가 없음!!)
```html
<!-- 텍스트 상자 -->
<div style='position: absolute; left: 100px; top: 50px; width: 200px; height: 100px; color: #000000; text-align: center;'>
    <span style='font-size: 24pt; font-weight: bold;'>제목</span>
</div>

<!-- 이미지 -->
<div style='position: absolute; left: 150px; top: 150px; width: 300px; height: 200px;'>
    <img src='{절대경로}/images/{GUID}.png' alt='Image' />
</div>

<!-- 도형 -->
<div style='position: absolute; left: 200px; top: 250px; width: 150px; height: 150px; background-color: rgba(255, 255, 255, 0.8); border-radius: 10px;'>
    <span style='font-size: 16pt;'>내용</span>
</div>
```

### 전체 슬라이드 HTML 구조(이 때는 각 슬라이드 페이지를 감싸는 div가 있음)
```html
<div class='Slide1'>
    <!-- 슬라이드 1의 내용 -->
</div>
<div class='Slide2'>
    <!-- 슬라이드 2의 내용 -->
</div>
<!-- 추가 슬라이드들... -->
```

### 슬라이드 요소 변환 예시
```html
<div class='Slide1'>
    <!-- 텍스트 상자 -->
    <div style='position: absolute; left: 100px; top: 50px; width: 200px; height: 100px; color: #000000; text-align: center;'>
        <span style='font-size: 24pt; font-weight: bold;'>제목</span>
    </div>
    
    <!-- 이미지 -->
    <div style='position: absolute; left: 150px; top: 150px; width: 300px; height: 200px;'>
        <img src='data:image/png;base64,...' alt='Image' />
    </div>
    
    <!-- 도형 -->
    <div style='position: absolute; left: 200px; top: 250px; width: 150px; height: 150px; background-color: rgba(255, 255, 255, 0.8); border-radius: 10px;'>
        <span style='font-size: 16pt;'>내용</span>
    </div>
</div>
```

### 저장 위치
- 변환된 HTML은 `test.html` 파일로 저장됩니다.
- 각 슬라이드의 모든 요소와 스타일이 보존됩니다.
- 슬라이드 번호는 `Slide1`, `Slide2` 등의 클래스로 구분됩니다. 
