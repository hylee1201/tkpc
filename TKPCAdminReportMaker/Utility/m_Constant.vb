Module m_Constant
    Public Const SQL_CONSTR As String = "Password=18deec12dafdb178db16d993a0d6970116de1ff760b8b2b795afbc862f9f82c3;Username=prnonoxpzhzqjp;Database=d4n80sith9p9lq;Host=ec2-184-73-199-189.compute-1.amazonaws.com;sslmode=Require;Trust Server Certificate=true"
    Public Const USER_ROLE_OFFICE_ADMIN As String = "office_admin"
    Public Const USER_ROLE_PASTROL_ADMIN As String = "pastoral_admin"
    Public Const ROOT_FOLDER As String = "C:\TKPC_REPORTS"
    Public Const PERSON_IMAGE_FOLDER As String = ROOT_FOLDER + "\images\person\"
    Public Const LOG_FOLDER As String = ROOT_FOLDER + "\logs\"
    Public Const FILE_FOLDER As String = ROOT_FOLDER + "\files\"
    Public Const TEMP_FILE As String = "temp.xls"
    Public Const FOUND_PEOPLE_FILE As String = "found_people_list.xls"
    Public Const MEMBER_FILE As String = "tkpc_member_list.xls"
    Public Const FAMILY_LEVEL_0 As String = "0-단독개인"
    Public Const FAMILY_LEVEL_1 As String = "1-세대주"
    Public Const FAMILY_LEVEL_2 As String = "2-부양가족"
    Public Const GENDER_M_KR As String = "남"
    Public Const GENDER_F_KR As String = "여"
    Public Const GENDER_M_EN As String = "M"
    Public Const GENDER_F_EN As String = "F"
    Public Const FAMILY_SPOUSE_EN As String = "spouse"
    Public Const FAMILY_SPOUSE_KR As String = "배우자"
    Public Const FAMILY_CHILD_KR As String = "자녀"
    Public Const FAMILY_ME_KR As String = "본인"
    Public Const FAMILY_TITLE As String = "가족"
    Public Const OFFERING_TOTAL As String = "Total"
    Public Const KbdKr As String = "00000412"  'Korean (standard)'//by default installed
    'Public Const KbdKr =  "00000012"  'Korean (standard)'//by default installed
    Public Const KbdEn As String = "00000409"  'English(US)   '//by default installed
    Public Const FORM_SURVEY As String = "tkpc_membership_survey.html"
    Public Const LABEL_SURVEY As String = "tkpc_membership_survey_label_5160.html"
    Public Const FAMILY_TYPE_HEAD As String = "세대주"
    Public Const FAMILY_TYPE_SPOUSE As String = "배우자"
    Public Const FAMILY_TYPE_KID As String = "자녀"

    Public Const ACCESS_LEVEL_PASTOR As String = "pastoral_admin"
    Public Const ACCESS_LEVEL_OFFICE As String = "office_admin"
    Public Const FILE_EXTENSION_LOG = ".log"
    Public Const FILE_EXTENSION_HTML = ".html"
    Public Const FILE_EXTENSION_PDF = ".pdf"

    Public Const FILE_MEMBERSHIP As Integer = 1
    Public Const FILE_SURVEY As Integer = 2

    Public Const PAPER_SIZE_LETTER As String = "letter"
    Public Const PAPER_ORIENTATION_PORTRAIT As String = "portrait"
    Public Const PAPER_ORIENTATION_LANDSCAPE As String = "landscape"
    Public Const PAPER_LETTER_LANDSCAPE As String = "letter landscape"
    Public Const PAPER_LETTER_PORTRAIT As String = "letter portrait"

    Public Const PROCESS_ACROBAT32 = "AcroRd32"
    Public Const PROCESS_ACROBAT = "Acrobat"

    Public Const RECEIPT_CHURCH_NAME = "church_name"
    Public Const RECEIPT_CHURCH_ADDR1 = "church_addr1"
    Public Const RECEIPT_CHURCH_ADDR2 = "church_addr2"
    Public Const RECEIPT_CHURCH_PHONE = "church_phone"
    Public Const RECEIPT_CHURCH_REG_NO = "church_reg_no"
    Public Const RECEIPT_TREASURER = "treasurer"

    Public Const MILESTONE_TITLE = "TitleMilestone"
    Public Const MILESTONE_BAPTISM = "BaptismMilestone"

    Public Const MILESTONE_NAME_SEORI = "서리집사"
    Public Const MILESTONE_NAME_ANSU = "안수집사"
    Public Const MILESTONE_NAME_SIMU_ELDER = "시무장로"
    Public Const MILESTONE_NAME_SIMU_KWONSA = "시무권사"

    Public Const PWD_OFFICE = "office_pwd"
    Public Const PWD_PASTOR = "pastor_pwd"
    Public Const QUESTION_OFFICE = "office_question"
    Public Const QUESTION_PASTOR = "pastor_question"
    Public Const ANSWER_OFFICE = "office_answer"
    Public Const ANSWER_PASTOR = "pastor_answer"

    Public Const MAX_OFFERING_NUMBER = "max_offering_number"

    Public Const MODE_VIEW_FAMILY = "가족 모드로 보기"
    Public Const MODE_VIEW_OFFERING = "전체 개인별"

    Public Const MEMBER_CHURCH = "church_member"
    Public Const MEMBER_LEFT = "left_church"
    Public Const MEMBER_NEW = "new_family"
    Public Const MEMBER_DECEASED = "deceased"

    Public Const AVERY_5160_MAX As Integer = 30
    Public Const ADDRESS_TKPC As String = "67 Scarsdale Rd, North York, Ontario, M3B2R2"

End Module
