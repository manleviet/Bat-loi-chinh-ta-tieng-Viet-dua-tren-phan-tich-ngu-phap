// Dự án:   BẮT LỖI CHÍNH TẢ TIẾNG VIỆT
// Tác giả: Lê Viết Mẫn
// Email:   manleviet@yahoo.com.hk

using System;

namespace perspicacity.VietnameseChecking.Model.Node
{
	/// <summary>
	/// Lớp _base là lớp cơ sở đầu tiên.
	/// Hỗ trợ việc xây dựng các từ điển.
	/// </summary>
	public class _base
	{
		#region Vùng khai báo biến
		/// <summary>
		/// Biến này nhận một chuỗi có thể là từ,
		/// từ loại, vế trái của luật sinh.
		/// </summary>
		private string _content1;

		/// <summary>
		/// Biến này nhận một chuỗi có thể là chuỗi từ loại,
		/// hay ý nghĩa của từ loại, hay vế phải của luật sinh.
		/// </summary>
		private string _content2;
		#endregion
		
		#region Vùng hàm khởi tạo
		/// <summary>
		/// Cấu tử cơ bản, không có tham số.
		/// </summary>
		public _base()
		{
			_content1 = "";
			_content2 = "";
		}

		/// <summary>
		/// Cấu tử thứ hai, có một tham số.
		/// </summary>
		/// <param name="content1"> 
		///	Nhận chuỗi có thể là từ, từ loại,
		///	vế trái của luật sinh.
		/// </param>
		/// <param name="content2"> 
		///	Nhận chuỗi có thể là chuỗi từ loại, hay ý nghĩa của từ loại,
		///	hay vế phải của luật sinh.
		/// </param>
		public _base(string content1, string content2)
		{
			_content1 = content1;
			_content2 = content2;
		}
		#endregion

		#region Vùng thuộc tính
		/// <summary>
		/// Thuộc tính Content1, lấy ra hay gán nội dung 
		/// của từ, từ loại, vế trái của luật sinh.
		/// </summary>
		protected string Content1
		{
			get
			{
				return _content1;
			}
			set
			{
				_content1 = value;
			}
		}

		/// <summary>
		/// Thuộc tính Content2, lấy ra hay gán nội dung của
		/// chuỗi từ loại, hay ý nghĩa từ loại, hay vế phải của luật sinh.
		/// </summary>
		protected string Content2
		{
			get
			{
				return _content2;
			}
			set
			{
				_content2 = value;
			}
		}
		#endregion
	}
}