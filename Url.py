import time


class Url:
    @staticmethod
    def get_url(url: str, params: dict = None):
        """url:str -> отформатированный url"""
        if (params is None) or (len(params) == 0):
            return url
        else:
            params_str = ""
            for key, value in params.items():
                params_str += key + "=" + value + "&"
            return url + "?" + params_str[:-1]

    @staticmethod
    def get_html(driver, url: str, params: dict = None):
        """html:str -> вернет странницу html"""
        driver.get(Url.get_url(url, params))
        print("Прогрузка страницы...")
        time.sleep(2)
        return driver.page_source
