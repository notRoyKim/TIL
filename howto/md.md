마크다운(markdown) 정리
=============


마크다운 사용법
---------------



# 마크다운 문법

## 1. header

+ 문서 제목
 
```
this is title
=============
```

* 결과
> this is title
> =============

+ 문서 부제

```
this is title
-------------
```

* 결과
> this is title
> -------------

+ 글머리 1 ~ 6


```
# this is an H1
## this is an H1
### this is an H1
#### this is an H1
##### this is an H1
###### this is an H1
```
* 결과
> # this is an H1
> ## this is an H1
> ### this is an H1
> #### this is an H1
> ##### this is an H1
> ###### this is an H1

## BlockQuote

```>```문자를 사용해 Blockquote 표시를 한다.

> this is a first blockquote
> > this is a second blockquote
> > > this is a third blockquote


## List

### 순서있는 목록
```
1. something
2. anything
3. everything
  3. else
```
(※ 위의 ```3. else```와 같이 작성해도 결과는 오름차순으로 숫자가 부여된다.)

1. something
2. anything
3. everything
  3. else

 ### 순서 없는 목록
 글머리 기호 ```*```, ```+```, ```-``` 사용
 
 ```
 * 1
 * 2
 * 3
 + 4
 - 5
  - 5-1
  + 5-2
    + 5-2-1
 ```
 
 * 1
 * 2
 * 3
 + 4
 - 5
   - 5-1
   + 5-2
     + 5-2-1
