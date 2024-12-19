# OpenPyXl_Based_Excel_to_Chroma_Vectorstore
This repository is showing to make about Excel sheets to Vectorstore of Chroma !
<br>


# 어떤 프로그램인가?
이 프로그램은 한 엑셀에 있는 차트에 대해서 Chroma의 '벡터스토어'로 변환하기 위해 Documents() 를 생성하는 방법을 증명하고 있습니다. 
다음과 같은 순서를 가집니다. ( 실제 구현을 위해 MITRE ATTACK 엑셀 파일을 샘플로 사용하였습니다. ) 

1. [Excel 절대경로 내놔!](https://github.com/lastime1650/OpenPyXl_Based_Excel_to_Chroma_Vectorstore/blob/f88b00891728526e4fcb79d7e22098d4aaffbe1b/Code/To_VectorStore.py#L29)
2. 클래스 생성자에 전달한다.
3. import된 Openpyxl 모듈에 기반으로 Excel 시트에 대한 정보를 WorkBook() 인스턴스형태료 변환받아 접근, 파싱가능
4. [파싱하여(꽤나 복잡함) 여러개의 Documents를 생성함 ( Page_Content, Metadata 도 추출할 수 있음 )](https://github.com/lastime1650/OpenPyXl_Based_Excel_to_Chroma_Vectorstore/blob/f88b00891728526e4fcb79d7e22098d4aaffbe1b/Code/To_VectorStore.py#L35)
5. [벡터스토어 인스턴스의 .add_documents()에 생성했던 여러개의 Documents를 추가하여 저장함!](https://github.com/lastime1650/OpenPyXl_Based_Excel_to_Chroma_Vectorstore/blob/f88b00891728526e4fcb79d7e22098d4aaffbe1b/Code/To_VectorStore.py#L138)
## 그러면 어떻게 생성된 벡터스토어를 사용하는가?..
여러 방법이 있지만, 이 코드에서는 Query한 문자열에 대해 "유사성"을 검사하고 반환하도록 하는 [.similarity_Search_with_relative_scores()](https://github.com/lastime1650/OpenPyXl_Based_Excel_to_Chroma_Vectorstore/blob/f88b00891728526e4fcb79d7e22098d4aaffbe1b/Code/To_VectorStore.py#L211) 를 호출합니다. 



<br>

# The following sample is here! ( Excel ) 
[MITRE_ATTACK_EXCEL](https://attack.mitre.org/resources/attack-data-and-tools)

위 페이지에 접근하여 MITRE 어택에 관한 엑셀파일을 다운로드 받을 수 있습니다. 

<br>

