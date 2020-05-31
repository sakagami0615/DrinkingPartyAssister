from assister import Assister 


if __name__ == '__main__':
	
	assist = Assister()
	
	assist.LoadWebPageTop()
	assist.SetShopConditions()

	shop_details = assist.GetShopDetail()
	course_details = assist.GetCourseDetail(shop_details)

	assist.AddResultSheet4Excel(course_details)

	assist.Close()