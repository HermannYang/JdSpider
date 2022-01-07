import json, re, time, requests, random, xlwt


class BaseSpider:

    def __init__(self):
        """
        初始化
        """
        self.comment_url = "https://club.jd.com/comment/productPageComments.action"
        # 定义请求头
        self.headers = {
            "referer": "https://item.jd.com/100006391096.html",     # 商品购买的URL,其中100006391096为3900x
            "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_0) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) C4hrome/81.0.404.138 Safari/537.36"}

    def get_comment(self, maxpage=100):
        """
        获取评论
        :param maxpage: 最大页数（0-100）
        :return:
        """
        data_list = []
        # 遍历列表  差评=1 中评=2 好评=3
        for _score in [1, 2, 3]:
            # 遍历所有页
            for page in range(maxpage):
                # 定义请求参数
                query = {
                    "callback": "fetchJSON_comment98",  # 默认
                    "productId": "100006391096",    # 商品ID
                    "score": _score,  # 差评1 中评2 好评3
                    "sortType": "5",    # 默认
                    "page": page,      # 当前页码
                    "pageSize": "10",   # 展示的评论数
                    "isShadowSku": "0",     # 默认
                    "fold": "1",    # 默认
                }
                # 设置延时
                time.sleep(random.randint(2, 5))
                res = requests.get(url=self.comment_url, headers=self.headers, params=query)
                res = json.loads(re.match(r"^fetchJSON_comment98\((.+)\);", res.text).group(1))
                for info in res["comments"]:
                    user_id = info["id"]    # 用户id
                    nickname = info["nickname"]     # 用户名称
                    content = info["content"].replace(" ", '').replace("\n", '')
                    creationTime = info["creationTime"] # 评论时间
                    referenceName = info["referenceName"]   # 商品名称
                    score = info["score"]   # 评论星数
                    info_dict = {
                        "page": page + 1,
                        "user_id": user_id,
                        "nickname": nickname,
                        "content": content,
                        "creationTime": creationTime,
                        "referenceName": referenceName,
                        "score": score,
                    }
                    if _score == 1:
                        info_dict["_score"] = "差评"
                        # 打印信息
                        print("【{_score}】页码：{page} 用户名：{nickname} 用户id：{user_id} "
                              "购买商品：{referenceName} 评论时间：{creationTime} "
                              "评论等级：{score}星 评论内容：{content}".format(**info_dict))
                    elif _score == 2:
                        info_dict["_score"] = "中评"
                        print("【{_score}】页码：{page} 用户名：{nickname} 用户id：{user_id} "
                              "购买商品：{referenceName} 评论时间：{creationTime} "
                              "评论等级：{score}星 评论内容：{content}".format(**info_dict))
                    else:
                        info_dict["_score"] = "好评"
                        print("【{_score}】页码：{page} 用户名：{nickname} 用户id：{user_id} "
                              "购买商品：{referenceName} 评论时间：{creationTime} "
                              "评论等级：{score}星 评论内容：{content}".format(**info_dict))
                    # 将字典添加到列表
                    data_list.append(info_dict)
        # 返回采集的数据
        return data_list

    def write_excle(self, data):
        """
        写入表
        :param data: 评论数据
        :return:
        """
        # 示例化
        book = xlwt.Workbook()
        sheet = book.add_sheet("sheet1")
        # 定义标题
        titles = ["用户id", "用户名", "购买商品", "页码", "评论时间", "评论类型", "评论等级", "评论内容"]
        for i, j in enumerate(titles):
            sheet.write(0, i, j)
        for i, j in enumerate(data):
            # 写入内容。
            sheet.write(i + 1, 0, j["user_id"])
            sheet.write(i + 1, 1, j["nickname"])
            sheet.write(i + 1, 2, j["referenceName"])
            sheet.write(i + 1, 3, j["page"])
            sheet.write(i + 1, 4, j["creationTime"])
            sheet.write(i + 1, 5, j["_score"])
            sheet.write(i + 1, 6, j["score"])
            sheet.write(i + 1, 7, j["content"])
        book.save("./Comment Data.xls")


if __name__ == '__main__':

    s = BaseSpider()
    data = s.get_comment()
    s.write_excle(data)
