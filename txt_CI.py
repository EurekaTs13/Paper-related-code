import pandas as pd

# Read Excel files
df1 = pd.read_excel(r"D:\excel\finance\topic_60_0.8_100_iter=5.xlsx")
filename_column_r = df1.columns[0]
filename_column_r_list = df1[filename_column_r].to_list()

df2 = pd.read_excel(r"D:\excel\Firm_SystemRisk_Finance-20230101_modified.xlsx")
filename_column = df2.columns[0]
filename_column_list = df2[filename_column].tolist()
year_column = df2.columns[2]
year_column_list = df2[year_column].tolist()

# Set the first column as the index
df1.set_index(df1.columns[0], inplace=True)

# Convert the DataFrame to a dictionary
data_dict = df1.apply(lambda row: row.values.tolist(), axis=1).to_dict()

# Write the dictionary to a txt file
with open(r"D:\excel\finance\out.txt", "w", encoding="utf-8") as file:
    for key, value in data_dict.items():
        file.write(f"{key}: {value}\n")

# Call the previously obtained similarity between topics
similarity = [[1.0, 0.4036, 0.372, 0.3507, 0.3272, 0.3828, 0.474, 0.3906, 0.6905, 0.3833, 0.3786, 0.433, 0.3698, 0.3132, 0.2739, 0.4185, 0.4146, 0.4475, 0.3536, 0.5067, 0.3928, 0.3895, 0.3157, 0.319, 0.4142, 0.297, 0.6293, 0.483, 0.3767, 0.4911, 0.3694, 0.5818, 0.4204, 0.3762, 0.4141, 0.4305, 0.3396, 0.4539, 0.3682, 0.4656, 0.3417, 0.321, 0.3353, 0.4004, 0.3529, 0.4493, 0.3632, 0.4764, 0.6663, 0.6413, 0.4905, 0.5571, 0.3623, 0.3432, 0.2783, 0.3502, 0.3694, 0.3169, 0.334, 0.3548, 0.3558, 0.4278, 0.4584, 0.4447, 0.4188, 0.4173, 0.3344, 0.4802, 0.3631, 0.3408, 0.3244, 0.3245, 0.3472, 0.4462, 0.4165, 0.3581, 0.3797, 0.3822, 0.3295, 0.4253], [0, 1.0, 0.4911, 0.6062, 0.5486, 0.3058, 0.5328, 0.5695, 0.5372, 0.8698, 0.4599, 0.4791, 0.38, 0.6256, 0.3429, 0.8783, 0.7851, 0.5186, 0.612, 0.793, 0.6169, 0.5163, 0.4048, 0.3096, 0.6064, 0.3876, 0.4981, 0.5074, 0.5615, 0.4847, 0.7957, 0.5064, 0.7926, 0.6593, 0.5037, 0.5563, 0.7557, 0.3785, 0.303, 0.5186, 0.6382, 0.7301, 0.5811, 0.4529, 0.4535, 0.4818, 0.7483, 0.5305, 0.4868, 0.4881, 0.3875, 0.4954, 0.5166, 0.5039, 0.4696, 0.6729, 0.5538, 0.4384, 0.5348, 0.5516, 0.4444, 0.7463, 0.4508, 0.6405, 0.5432, 0.4916, 0.3231, 0.4046, 0.4616, 0.5292, 0.5505, 0.4808, 0.6155, 0.5401, 0.7534, 0.555, 0.4545, 0.3285, 0.6124, 0.6759], [0, 0, 1.0, 0.6527, 0.6728, 0.3687, 0.5972, 0.5961, 0.4175, 0.5764, 0.504, 0.5454, 0.4907, 0.6427, 0.4377, 0.5331, 0.6231, 0.6692, 0.5244, 0.4849, 0.5784, 0.6106, 0.5257, 0.5004, 0.6426, 0.5449, 0.5461, 0.6398, 0.627, 0.5203, 0.5615, 0.5628, 0.542, 0.5361, 0.5816, 0.636, 0.5729, 0.4461, 0.3683, 0.5383, 0.5829, 0.4587, 0.6355, 0.5422, 0.6191, 0.5712, 0.5741, 0.6725, 0.4037, 0.4284, 0.4408, 0.4816, 0.5962, 0.5902, 0.5989, 0.528, 0.628, 0.6465, 0.7158, 0.5684, 0.5995, 0.564, 0.5594, 0.4978, 0.6277, 0.5863, 0.4275, 0.4969, 0.5833, 0.6818, 0.46, 0.6384, 0.5629, 0.5634, 0.5795, 0.6893, 0.463, 0.347, 0.5885, 0.5169], [0, 0, 0, 1.0, 0.7018, 0.3721, 0.5981, 0.5744, 0.4119, 0.6837, 0.4289, 0.5789, 0.395, 0.7027, 0.4332, 0.6152, 0.6977, 0.6209, 0.571, 0.5648, 0.563, 0.5962, 0.4808, 0.4418, 0.5817, 0.4348, 0.5514, 0.5376, 0.6072, 0.4781, 0.6995, 0.5329, 0.6423, 0.6195, 0.6538, 0.6303, 0.6256, 0.3406, 0.3445, 0.4681, 0.6751, 0.571, 0.6018, 0.5072, 0.6239, 0.5264, 0.6579, 0.5891, 0.408, 0.4537, 0.3445, 0.4331, 0.5415, 0.6679, 0.4824, 0.6628, 0.6023, 0.5641, 0.6338, 0.5504, 0.5172, 0.5884, 0.4536, 0.5552, 0.59, 0.5257, 0.3842, 0.4552, 0.6241, 0.6888, 0.53, 0.6102, 0.6832, 0.5394, 0.6231, 0.6606, 0.4446, 0.33, 0.6122, 0.5306], [0, 0, 0, 0, 1.0, 0.3947, 0.6001, 0.5503, 0.4086, 0.637, 0.5329, 0.5304, 0.4957, 0.7138, 0.4328, 0.5785, 0.6551, 0.6337, 0.5969, 0.5158, 0.6071, 0.6059, 0.5478, 0.4984, 0.6213, 0.5472, 0.6303, 0.5577, 0.5961, 0.5193, 0.5884, 0.5685, 0.5563, 0.6196, 0.5483, 0.7001, 0.6237, 0.4654, 0.3745, 0.5286, 0.6216, 0.4847, 0.7246, 0.5076, 0.6597, 0.5622, 0.5736, 0.6256, 0.3716, 0.4059, 0.4185, 0.5019, 0.6754, 0.5521, 0.5338, 0.592, 0.7125, 0.5957, 0.6739, 0.6104, 0.5822, 0.5753, 0.5656, 0.5771, 0.6861, 0.5707, 0.3902, 0.4651, 0.5815, 0.7332, 0.5151, 0.6346, 0.6455, 0.552, 0.5751, 0.6739, 0.4892, 0.3397, 0.7221, 0.5111], [0, 0, 0, 0, 0, 1.0, 0.4301, 0.4315, 0.2883, 0.354, 0.3737, 0.429, 0.396, 0.3778, 0.4772, 0.3394, 0.41, 0.4157, 0.4536, 0.3314, 0.3025, 0.6768, 0.6616, 0.466, 0.4559, 0.411, 0.3765, 0.4816, 0.3585, 0.4742, 0.3485, 0.4012, 0.3531, 0.3704, 0.3888, 0.4365, 0.342, 0.374, 0.525, 0.4314, 0.3789, 0.3023, 0.3849, 0.6086, 0.4671, 0.6976, 0.4239, 0.4531, 0.3549, 0.3133, 0.4788, 0.3584, 0.3336, 0.4491, 0.2581, 0.3155, 0.3847, 0.4238, 0.3429, 0.4748, 0.4663, 0.4175, 0.3898, 0.3338, 0.452, 0.4124, 0.3949, 0.4062, 0.3774, 0.3344, 0.354, 0.3889, 0.3505, 0.66, 0.3673, 0.3713, 0.7466, 0.3627, 0.3905, 0.3722], [0, 0, 0, 0, 0, 0, 1.0, 0.6217, 0.5414, 0.5637, 0.549, 0.5592, 0.6056, 0.5979, 0.5773, 0.5167, 0.5774, 0.6182, 0.509, 0.5543, 0.548, 0.5584, 0.5265, 0.4786, 0.6122, 0.562, 0.5565, 0.6327, 0.5229, 0.6385, 0.5296, 0.5027, 0.5851, 0.5628, 0.5201, 0.6354, 0.5521, 0.5119, 0.4886, 0.5379, 0.5302, 0.4525, 0.5726, 0.5637, 0.5895, 0.6352, 0.5705, 0.594, 0.5135, 0.45, 0.5534, 0.6392, 0.542, 0.5695, 0.4767, 0.5588, 0.617, 0.5495, 0.5662, 0.5057, 0.6089, 0.5197, 0.5739, 0.5391, 0.5976, 0.6425, 0.5209, 0.6975, 0.6292, 0.611, 0.413, 0.5143, 0.556, 0.5577, 0.5876, 0.5556, 0.5081, 0.5363, 0.5569, 0.5458], [0, 0, 0, 0, 0, 0, 0, 1.0, 0.4529, 0.6182, 0.4831, 0.5425, 0.4918, 0.567, 0.6033, 0.5809, 0.6196, 0.5955, 0.5137, 0.5568, 0.5518, 0.5883, 0.5397, 0.4993, 0.5921, 0.5206, 0.5215, 0.6307, 0.5718, 0.5558, 0.6325, 0.4624, 0.6494, 0.6122, 0.5154, 0.5981, 0.602, 0.5043, 0.4861, 0.5304, 0.5647, 0.5379, 0.5707, 0.5108, 0.5506, 0.6279, 0.6555, 0.629, 0.4819, 0.4098, 0.4815, 0.5091, 0.5871, 0.6338, 0.4298, 0.565, 0.6028, 0.5329, 0.5459, 0.5569, 0.6714, 0.5912, 0.5032, 0.5651, 0.5871, 0.5845, 0.5556, 0.5674, 0.5529, 0.5892, 0.4522, 0.515, 0.5652, 0.5904, 0.6147, 0.5565, 0.5066, 0.4013, 0.6049, 0.5893], [0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4869, 0.439, 0.4562, 0.3571, 0.4067, 0.3199, 0.4775, 0.4961, 0.4569, 0.4491, 0.5771, 0.4959, 0.3948, 0.3107, 0.3228, 0.4634, 0.3402, 0.6732, 0.5245, 0.4774, 0.5011, 0.449, 0.5341, 0.4954, 0.471, 0.4107, 0.4888, 0.4634, 0.4529, 0.3627, 0.51, 0.4507, 0.4305, 0.4265, 0.4697, 0.4132, 0.4658, 0.4439, 0.488, 0.6055, 0.669, 0.4945, 0.6223, 0.4092, 0.382, 0.3402, 0.4576, 0.4586, 0.3676, 0.4106, 0.4195, 0.4267, 0.4537, 0.4496, 0.4823, 0.488, 0.462, 0.3304, 0.5139, 0.4434, 0.4232, 0.3791, 0.3622, 0.4705, 0.4443, 0.4977, 0.4059, 0.3912, 0.3406, 0.4385, 0.4833], [0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5178, 0.5226, 0.4431, 0.718, 0.4352, 0.9094, 0.8232, 0.554, 0.6991, 0.7871, 0.6765, 0.6057, 0.4938, 0.3483, 0.5916, 0.4844, 0.5491, 0.5553, 0.6308, 0.4916, 0.8853, 0.5357, 0.8686, 0.6963, 0.5665, 0.629, 0.8568, 0.4007, 0.3409, 0.5245, 0.7164, 0.7871, 0.662, 0.4467, 0.5464, 0.5527, 0.792, 0.579, 0.4293, 0.4957, 0.382, 0.4622, 0.539, 0.5794, 0.5258, 0.6884, 0.6528, 0.575, 0.6367, 0.6399, 0.5251, 0.8063, 0.4781, 0.6601, 0.5859, 0.574, 0.3959, 0.4169, 0.4777, 0.6169, 0.6421, 0.6013, 0.6595, 0.5873, 0.8393, 0.6626, 0.5051, 0.3691, 0.6559, 0.7158], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4093, 0.6718, 0.5033, 0.3925, 0.4438, 0.4922, 0.5046, 0.4408, 0.4531, 0.5774, 0.5708, 0.4614, 0.4149, 0.5008, 0.4584, 0.51, 0.5625, 0.4376, 0.5029, 0.4447, 0.4758, 0.4887, 0.4872, 0.4276, 0.6114, 0.497, 0.7898, 0.2955, 0.421, 0.4274, 0.431, 0.5087, 0.3926, 0.462, 0.5797, 0.4304, 0.5578, 0.3672, 0.3685, 0.4045, 0.596, 0.4535, 0.3915, 0.4181, 0.4135, 0.5939, 0.4999, 0.4774, 0.4915, 0.5104, 0.4676, 0.4206, 0.5852, 0.4733, 0.683, 0.4222, 0.4615, 0.4559, 0.5542, 0.3608, 0.4655, 0.4914, 0.5988, 0.5334, 0.5002, 0.5519, 0.4552, 0.4542, 0.4925], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4235, 0.5452, 0.4959, 0.453, 0.5634, 0.5322, 0.5364, 0.418, 0.4472, 0.4711, 0.3835, 0.5681, 0.5584, 0.5054, 0.5629, 0.517, 0.4371, 0.4705, 0.4912, 0.5537, 0.437, 0.4233, 0.5011, 0.4811, 0.5017, 0.3008, 0.4978, 0.386, 0.6147, 0.3388, 0.4738, 0.5304, 0.5625, 0.4648, 0.4567, 0.4947, 0.4816, 0.4727, 0.4307, 0.4479, 0.5177, 0.5628, 0.3934, 0.4682, 0.4995, 0.4322, 0.5428, 0.4137, 0.5367, 0.4425, 0.5178, 0.4162, 0.5134, 0.4258, 0.3871, 0.5248, 0.5426, 0.5255, 0.4002, 0.4606, 0.5287, 0.4701, 0.4454, 0.5034, 0.3912, 0.45, 0.5636, 0.4099], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4877, 0.4894, 0.3717, 0.4287, 0.4975, 0.389, 0.3813, 0.4838, 0.5035, 0.4415, 0.5597, 0.4716, 0.5364, 0.401, 0.5124, 0.4533, 0.446, 0.3871, 0.3758, 0.4495, 0.4226, 0.457, 0.5941, 0.4438, 0.5663, 0.3329, 0.417, 0.3802, 0.3685, 0.545, 0.4359, 0.57, 0.5284, 0.4426, 0.4894, 0.389, 0.2989, 0.4122, 0.4921, 0.4793, 0.4163, 0.3682, 0.3704, 0.5322, 0.5299, 0.4945, 0.4291, 0.5128, 0.443, 0.4643, 0.4921, 0.4628, 0.6711, 0.4199, 0.5248, 0.4824, 0.4897, 0.3118, 0.4655, 0.4657, 0.5003, 0.5134, 0.5094, 0.4372, 0.4463, 0.4895, 0.4611], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4678, 0.5807, 0.628, 0.652, 0.5435, 0.5386, 0.581, 0.5799, 0.4957, 0.4691, 0.6043, 0.5341, 0.5512, 0.5745, 0.6085, 0.468, 0.6056, 0.524, 0.5958, 0.5909, 0.5545, 0.6811, 0.6811, 0.4405, 0.3977, 0.4931, 0.6329, 0.5459, 0.6511, 0.4914, 0.6453, 0.5272, 0.5557, 0.6087, 0.3924, 0.4256, 0.3719, 0.4655, 0.6088, 0.5573, 0.4882, 0.6029, 0.7245, 0.5045, 0.6288, 0.5388, 0.5577, 0.5673, 0.4934, 0.5315, 0.6082, 0.5741, 0.386, 0.5021, 0.5335, 0.7291, 0.4823, 0.5844, 0.6818, 0.5086, 0.6014, 0.8587, 0.4543, 0.4317, 0.6162, 0.5093], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.3217, 0.3913, 0.4317, 0.3754, 0.346, 0.3816, 0.4583, 0.4676, 0.5999, 0.511, 0.7402, 0.3682, 0.569, 0.4063, 0.6062, 0.3972, 0.3724, 0.4312, 0.41, 0.3815, 0.507, 0.44, 0.3919, 0.6794, 0.3426, 0.4637, 0.3461, 0.4483, 0.5089, 0.5572, 0.5119, 0.4573, 0.4029, 0.3908, 0.2674, 0.5713, 0.4606, 0.3952, 0.5952, 0.3189, 0.4284, 0.4564, 0.5231, 0.4623, 0.387, 0.7237, 0.3878, 0.5612, 0.3943, 0.5337, 0.5475, 0.6306, 0.6086, 0.4613, 0.4107, 0.2779, 0.3601, 0.456, 0.457, 0.4097, 0.3742, 0.4005, 0.5144, 0.4838, 0.4103], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.808, 0.5476, 0.6464, 0.861, 0.6807, 0.5706, 0.4538, 0.301, 0.5927, 0.3431, 0.5292, 0.5208, 0.5786, 0.5057, 0.8384, 0.5234, 0.8629, 0.6942, 0.5178, 0.5884, 0.774, 0.3573, 0.2901, 0.5694, 0.608, 0.809, 0.5771, 0.4144, 0.4736, 0.5392, 0.8013, 0.6167, 0.4214, 0.5228, 0.3522, 0.4713, 0.5218, 0.5434, 0.5247, 0.6903, 0.5762, 0.5152, 0.604, 0.5703, 0.4279, 0.8704, 0.4531, 0.6869, 0.5259, 0.5382, 0.3404, 0.3699, 0.4636, 0.5485, 0.5948, 0.5592, 0.5617, 0.5958, 0.8323, 0.5613, 0.5301, 0.2799, 0.6004, 0.7483], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.609, 0.664, 0.6932, 0.6354, 0.6065, 0.5312, 0.3764, 0.5852, 0.4767, 0.5688, 0.581, 0.5902, 0.4909, 0.8192, 0.553, 0.758, 0.6171, 0.5698, 0.6037, 0.7114, 0.4042, 0.3184, 0.571, 0.6896, 0.5991, 0.6593, 0.5245, 0.5544, 0.5588, 0.8177, 0.6423, 0.4534, 0.4743, 0.3944, 0.4805, 0.5148, 0.5806, 0.5514, 0.5834, 0.6075, 0.5603, 0.6142, 0.6399, 0.5264, 0.7699, 0.4507, 0.5806, 0.6201, 0.5467, 0.3547, 0.4657, 0.5399, 0.6274, 0.6099, 0.6276, 0.5809, 0.6173, 0.701, 0.5993, 0.5333, 0.312, 0.638, 0.6104], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4936, 0.5205, 0.5711, 0.6126, 0.5179, 0.5827, 0.5604, 0.429, 0.5841, 0.5969, 0.4821, 0.4928, 0.5235, 0.4968, 0.5529, 0.5164, 0.4622, 0.6603, 0.5204, 0.481, 0.3815, 0.5495, 0.4587, 0.4422, 0.5488, 0.5333, 0.6123, 0.5834, 0.5589, 0.7378, 0.3904, 0.4724, 0.3352, 0.5004, 0.5555, 0.5677, 0.5593, 0.531, 0.6805, 0.5111, 0.5765, 0.5024, 0.4798, 0.5905, 0.4801, 0.5225, 0.5456, 0.5814, 0.3915, 0.5266, 0.6563, 0.7574, 0.3838, 0.641, 0.5176, 0.6148, 0.5908, 0.6893, 0.5144, 0.3648, 0.5472, 0.5326], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5749, 0.5469, 0.6507, 0.6171, 0.379, 0.5446, 0.4399, 0.5242, 0.5288, 0.5553, 0.4774, 0.6464, 0.4841, 0.6001, 0.5704, 0.5661, 0.5182, 0.6359, 0.3431, 0.3439, 0.4756, 0.7334, 0.5512, 0.5855, 0.4906, 0.5614, 0.5928, 0.5999, 0.5395, 0.3865, 0.4091, 0.3973, 0.3758, 0.481, 0.5747, 0.4545, 0.5832, 0.5627, 0.5129, 0.5892, 0.812, 0.5206, 0.5878, 0.4462, 0.5514, 0.5821, 0.4666, 0.3889, 0.3853, 0.5045, 0.5234, 0.7891, 0.6198, 0.6329, 0.6017, 0.5978, 0.5409, 0.5703, 0.3405, 0.7027, 0.5656], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6829, 0.569, 0.4284, 0.3206, 0.5792, 0.3751, 0.5967, 0.5225, 0.5437, 0.5756, 0.7366, 0.5346, 0.8295, 0.6943, 0.4478, 0.6341, 0.7151, 0.4284, 0.3259, 0.6368, 0.5407, 0.8019, 0.5499, 0.433, 0.4981, 0.5534, 0.7162, 0.6584, 0.461, 0.6024, 0.382, 0.6593, 0.5471, 0.516, 0.4457, 0.7289, 0.6027, 0.5073, 0.5713, 0.5368, 0.4661, 0.7928, 0.5242, 0.7418, 0.4902, 0.5614, 0.3637, 0.4336, 0.4675, 0.5069, 0.5333, 0.5031, 0.5309, 0.6337, 0.796, 0.5222, 0.5688, 0.3114, 0.5606, 0.7225], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6042, 0.3978, 0.4727, 0.6428, 0.413, 0.6229, 0.5611, 0.5596, 0.6155, 0.5913, 0.5698, 0.6534, 0.7724, 0.4603, 0.7247, 0.6858, 0.4152, 0.3122, 0.5246, 0.4815, 0.6962, 0.6686, 0.3867, 0.5646, 0.6034, 0.5483, 0.6611, 0.3978, 0.466, 0.4729, 0.5774, 0.6229, 0.4871, 0.4877, 0.6726, 0.7014, 0.5417, 0.6272, 0.5262, 0.5489, 0.6285, 0.5291, 0.8308, 0.5765, 0.6224, 0.393, 0.3731, 0.5213, 0.6828, 0.4859, 0.54, 0.5091, 0.6188, 0.7082, 0.598, 0.5717, 0.3112, 0.6395, 0.7016], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.8061, 0.5041, 0.6848, 0.4971, 0.5305, 0.6915, 0.5212, 0.6124, 0.5917, 0.5165, 0.5893, 0.582, 0.5501, 0.6626, 0.6143, 0.4888, 0.3852, 0.6127, 0.5908, 0.5407, 0.633, 0.5372, 0.6297, 0.877, 0.5836, 0.721, 0.3756, 0.4356, 0.4671, 0.5207, 0.5806, 0.6005, 0.4462, 0.5492, 0.625, 0.5484, 0.5686, 0.6949, 0.6564, 0.5654, 0.5024, 0.6476, 0.5319, 0.5707, 0.5011, 0.4735, 0.5467, 0.609, 0.5653, 0.6162, 0.5485, 0.9203, 0.6034, 0.5927, 0.8532, 0.3943, 0.6367, 0.5352], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.3712, 0.5148, 0.527, 0.3738, 0.598, 0.4395, 0.4972, 0.4939, 0.378, 0.4966, 0.4126, 0.4888, 0.5155, 0.4926, 0.487, 0.3331, 0.5745, 0.5423, 0.3649, 0.543, 0.4971, 0.5283, 0.7526, 0.5397, 0.5291, 0.34, 0.3, 0.3603, 0.3866, 0.4123, 0.521, 0.4341, 0.4201, 0.4703, 0.5831, 0.519, 0.6904, 0.604, 0.5207, 0.4511, 0.438, 0.519, 0.5425, 0.4995, 0.4134, 0.4885, 0.4911, 0.5482, 0.6555, 0.474, 0.7633, 0.4599, 0.4958, 0.7637, 0.3589, 0.5296, 0.411], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5151, 0.5037, 0.4295, 0.4713, 0.3827, 0.4771, 0.3277, 0.3831, 0.3566, 0.4706, 0.3838, 0.5621, 0.4172, 0.314, 0.5819, 0.3489, 0.3877, 0.3412, 0.4817, 0.5625, 0.7089, 0.4758, 0.372, 0.5074, 0.3084, 0.2969, 0.4239, 0.4039, 0.4855, 0.6038, 0.3197, 0.4122, 0.5356, 0.4282, 0.4886, 0.3902, 0.5692, 0.3671, 0.517, 0.4652, 0.4837, 0.4704, 0.414, 0.4774, 0.5131, 0.5161, 0.2897, 0.3779, 0.4334, 0.4474, 0.3991, 0.4512, 0.3879, 0.4136, 0.511, 0.4272], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.567, 0.5402, 0.6582, 0.5175, 0.7083, 0.5313, 0.5322, 0.5392, 0.6141, 0.4911, 0.6325, 0.6077, 0.4105, 0.4861, 0.5752, 0.612, 0.5061, 0.6589, 0.5547, 0.6056, 0.6647, 0.5489, 0.6614, 0.4608, 0.4431, 0.5959, 0.6003, 0.6991, 0.5617, 0.436, 0.5718, 0.5602, 0.4864, 0.587, 0.5353, 0.7202, 0.5398, 0.5663, 0.6332, 0.5587, 0.566, 0.4809, 0.5454, 0.6044, 0.5983, 0.4374, 0.4594, 0.5544, 0.6678, 0.5418, 0.514, 0.6207, 0.437, 0.7038, 0.5551], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4149, 0.5397, 0.4803, 0.5366, 0.4345, 0.4389, 0.437, 0.414, 0.4499, 0.564, 0.5194, 0.4445, 0.5067, 0.3868, 0.566, 0.3183, 0.5957, 0.5459, 0.6009, 0.5119, 0.4959, 0.4151, 0.3704, 0.2568, 0.5507, 0.5183, 0.4774, 0.5503, 0.4023, 0.4331, 0.5376, 0.6245, 0.572, 0.492, 0.7552, 0.4276, 0.6315, 0.4324, 0.568, 0.5826, 0.5076, 0.5588, 0.4546, 0.4936, 0.3468, 0.4878, 0.5309, 0.4814, 0.4421, 0.4263, 0.3989, 0.4992, 0.6133, 0.4048], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5732, 0.5197, 0.5992, 0.5098, 0.7669, 0.5034, 0.5616, 0.4886, 0.6904, 0.51, 0.4388, 0.3756, 0.5082, 0.5015, 0.4807, 0.5493, 0.4798, 0.5219, 0.5643, 0.4646, 0.6678, 0.5378, 0.6997, 0.5529, 0.6305, 0.5748, 0.5063, 0.3697, 0.5633, 0.7235, 0.4411, 0.5192, 0.4774, 0.531, 0.5053, 0.6456, 0.6125, 0.5431, 0.4933, 0.4208, 0.4749, 0.4971, 0.619, 0.4524, 0.4979, 0.534, 0.5192, 0.5356, 0.6061, 0.4937, 0.3198, 0.5127, 0.512], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4937, 0.6667, 0.5097, 0.5832, 0.5455, 0.5276, 0.5262, 0.6388, 0.5459, 0.5462, 0.4638, 0.5661, 0.5359, 0.4514, 0.5576, 0.5347, 0.5862, 0.7251, 0.5456, 0.6942, 0.5479, 0.4813, 0.5628, 0.6068, 0.5715, 0.5595, 0.4275, 0.4963, 0.5735, 0.5098, 0.5408, 0.543, 0.6824, 0.535, 0.5344, 0.5681, 0.5746, 0.6117, 0.543, 0.603, 0.5747, 0.5588, 0.4173, 0.5189, 0.5642, 0.6952, 0.5526, 0.5547, 0.614, 0.4612, 0.5325, 0.4801], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.493, 0.6751, 0.5435, 0.5854, 0.578, 0.5728, 0.581, 0.5966, 0.4122, 0.3697, 0.3972, 0.6114, 0.5919, 0.6671, 0.481, 0.5735, 0.5415, 0.6354, 0.5392, 0.4766, 0.4052, 0.4774, 0.4537, 0.4867, 0.6163, 0.4772, 0.6237, 0.579, 0.6449, 0.6589, 0.5843, 0.5123, 0.5721, 0.5128, 0.5181, 0.606, 0.5047, 0.4075, 0.4195, 0.4901, 0.5558, 0.5347, 0.5156, 0.7181, 0.4639, 0.6005, 0.6832, 0.4081, 0.3318, 0.6356, 0.5559], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4789, 0.5646, 0.5198, 0.5554, 0.4561, 0.642, 0.5112, 0.4775, 0.6174, 0.4977, 0.497, 0.4896, 0.5412, 0.4954, 0.5373, 0.66, 0.5161, 0.6377, 0.506, 0.4661, 0.669, 0.7118, 0.5733, 0.5646, 0.4382, 0.5531, 0.5555, 0.5515, 0.5541, 0.5079, 0.7018, 0.4708, 0.6886, 0.6525, 0.5325, 0.5628, 0.7222, 0.5952, 0.5102, 0.4932, 0.4145, 0.4461, 0.496, 0.6241, 0.5242, 0.4597, 0.5942, 0.4766, 0.4978, 0.544], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4832, 0.8713, 0.648, 0.5711, 0.5712, 0.7645, 0.366, 0.3319, 0.4958, 0.7004, 0.7666, 0.6213, 0.4299, 0.5098, 0.524, 0.8817, 0.5493, 0.407, 0.4845, 0.3529, 0.4172, 0.4805, 0.6847, 0.5349, 0.6932, 0.5499, 0.6015, 0.6178, 0.6302, 0.5046, 0.7744, 0.4646, 0.6119, 0.5824, 0.4857, 0.4098, 0.4116, 0.4714, 0.5628, 0.644, 0.5766, 0.6761, 0.577, 0.7941, 0.5888, 0.4681, 0.3358, 0.6312, 0.6744], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4588, 0.4768, 0.4953, 0.6049, 0.4608, 0.4094, 0.4041, 0.4442, 0.5447, 0.4286, 0.5403, 0.4504, 0.498, 0.5346, 0.4668, 0.5845, 0.5487, 0.6893, 0.5968, 0.5514, 0.4851, 0.4712, 0.3623, 0.4866, 0.5823, 0.4562, 0.4954, 0.4604, 0.5103, 0.5176, 0.6186, 0.5262, 0.5357, 0.4972, 0.4141, 0.4385, 0.4057, 0.5221, 0.4508, 0.4405, 0.5315, 0.5177, 0.4813, 0.5631, 0.4701, 0.4055, 0.49, 0.475], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.7193, 0.5138, 0.6119, 0.8026, 0.4335, 0.3439, 0.5625, 0.6066, 0.8367, 0.5553, 0.4323, 0.5269, 0.5567, 0.828, 0.5683, 0.42, 0.4897, 0.3776, 0.5334, 0.5136, 0.6006, 0.5235, 0.7077, 0.5915, 0.6111, 0.6106, 0.599, 0.5094, 0.8274, 0.4554, 0.7112, 0.5456, 0.6006, 0.4268, 0.479, 0.4861, 0.5612, 0.6084, 0.5485, 0.6005, 0.6211, 0.876, 0.5554, 0.5096, 0.3704, 0.5816, 0.7813], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4879, 0.6744, 0.7744, 0.3879, 0.3422, 0.4415, 0.5406, 0.8017, 0.5931, 0.4391, 0.6175, 0.5891, 0.5986, 0.5419, 0.3966, 0.4422, 0.4899, 0.4918, 0.641, 0.5722, 0.452, 0.7643, 0.6867, 0.5285, 0.6307, 0.535, 0.5328, 0.6274, 0.495, 0.7891, 0.5425, 0.572, 0.3949, 0.4179, 0.5255, 0.6184, 0.5334, 0.5026, 0.5898, 0.5669, 0.7593, 0.5613, 0.5195, 0.3152, 0.6288, 0.7528], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.56, 0.5309, 0.3858, 0.3114, 0.4208, 0.5912, 0.4331, 0.5752, 0.4621, 0.6172, 0.5337, 0.5507, 0.5125, 0.4834, 0.3984, 0.3693, 0.3654, 0.4226, 0.6346, 0.465, 0.5066, 0.5157, 0.5564, 0.617, 0.566, 0.483, 0.491, 0.4921, 0.4898, 0.5762, 0.4731, 0.3672, 0.3421, 0.5058, 0.4771, 0.4795, 0.5507, 0.6038, 0.4807, 0.5273, 0.6164, 0.4114, 0.3397, 0.5496, 0.4523], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6077, 0.5394, 0.4185, 0.5636, 0.5038, 0.6088, 0.7037, 0.5167, 0.6345, 0.6553, 0.5819, 0.7336, 0.4081, 0.4427, 0.4987, 0.6766, 0.6211, 0.605, 0.4815, 0.6697, 0.8398, 0.627, 0.6912, 0.559, 0.6206, 0.613, 0.6907, 0.7836, 0.5801, 0.6896, 0.4693, 0.5168, 0.5891, 0.6835, 0.4302, 0.6103, 0.5963, 0.645, 0.6711, 0.713, 0.6082, 0.3955, 0.6489, 0.6009], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.3712, 0.3404, 0.504, 0.715, 0.7901, 0.6543, 0.4283, 0.5629, 0.5494, 0.7022, 0.5343, 0.3768, 0.4428, 0.3799, 0.4849, 0.6123, 0.5715, 0.5089, 0.7075, 0.6474, 0.5736, 0.6789, 0.6326, 0.5771, 0.6912, 0.4693, 0.6983, 0.5646, 0.5616, 0.4507, 0.4143, 0.5161, 0.6333, 0.6433, 0.5162, 0.6186, 0.5613, 0.8208, 0.5842, 0.4902, 0.3687, 0.646, 0.7235], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.3241, 0.4557, 0.3618, 0.3417, 0.4352, 0.4138, 0.3936, 0.5354, 0.3911, 0.5216, 0.4258, 0.3647, 0.3882, 0.6116, 0.3855, 0.3513, 0.3832, 0.3646, 0.4999, 0.481, 0.4307, 0.4421, 0.4578, 0.4287, 0.3866, 0.4599, 0.4816, 0.6631, 0.4008, 0.4976, 0.4348, 0.4444, 0.3141, 0.4708, 0.4585, 0.5447, 0.4216, 0.4465, 0.4803, 0.4743, 0.3799, 0.4292], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.3338, 0.3928, 0.3102, 0.3927, 0.5235, 0.4536, 0.4568, 0.3836, 0.392, 0.3979, 0.3118, 0.6148, 0.4013, 0.3475, 0.4633, 0.3201, 0.3557, 0.3728, 0.4273, 0.3881, 0.3429, 0.5617, 0.3698, 0.5196, 0.3001, 0.4988, 0.3465, 0.4963, 0.557, 0.3411, 0.3236, 0.2701, 0.2851, 0.3538, 0.3762, 0.3514, 0.3349, 0.3604, 0.5153, 0.405, 0.388], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.444, 0.4686, 0.501, 0.4847, 0.473, 0.5742, 0.5436, 0.6814, 0.3774, 0.452, 0.34, 0.5553, 0.4897, 0.4216, 0.407, 0.4968, 0.5036, 0.454, 0.4907, 0.4803, 0.4715, 0.5683, 0.4513, 0.509, 0.5149, 0.5012, 0.3812, 0.4248, 0.4915, 0.5056, 0.3857, 0.5259, 0.4453, 0.6358, 0.5926, 0.4808, 0.6037, 0.3295, 0.4879, 0.4812], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5143, 0.6263, 0.4867, 0.5718, 0.4972, 0.6511, 0.4717, 0.4503, 0.4241, 0.4508, 0.41, 0.5076, 0.6804, 0.4603, 0.6129, 0.5095, 0.5345, 0.6209, 0.7077, 0.6465, 0.5595, 0.5014, 0.4958, 0.5971, 0.4646, 0.4334, 0.4138, 0.5446, 0.5536, 0.7572, 0.5184, 0.7392, 0.5287, 0.5757, 0.5471, 0.427, 0.3872, 0.6953, 0.4936], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5247, 0.3436, 0.4812, 0.5161, 0.6869, 0.5155, 0.3427, 0.4737, 0.3589, 0.4458, 0.499, 0.5528, 0.423, 0.8076, 0.5924, 0.5146, 0.6143, 0.4848, 0.4165, 0.733, 0.4534, 0.7588, 0.4775, 0.5233, 0.3615, 0.3418, 0.3853, 0.475, 0.548, 0.4727, 0.5795, 0.5242, 0.853, 0.5194, 0.4904, 0.3026, 0.5365, 0.8121], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4769, 0.6499, 0.5844, 0.6083, 0.5988, 0.3972, 0.3679, 0.4473, 0.4698, 0.6088, 0.5727, 0.5256, 0.5937, 0.694, 0.6795, 0.7584, 0.6619, 0.6185, 0.5744, 0.5498, 0.6088, 0.6638, 0.5993, 0.4131, 0.4213, 0.5356, 0.6701, 0.5268, 0.6663, 0.607, 0.5705, 0.6086, 0.6669, 0.5015, 0.336, 0.707, 0.5357], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5632, 0.5564, 0.5009, 0.5016, 0.4051, 0.383, 0.4729, 0.4962, 0.4611, 0.524, 0.4013, 0.4425, 0.4687, 0.4416, 0.4525, 0.489, 0.5189, 0.5073, 0.4974, 0.4031, 0.5627, 0.4945, 0.4211, 0.5581, 0.5317, 0.4716, 0.3621, 0.4217, 0.465, 0.5088, 0.4436, 0.4698, 0.4821, 0.4125, 0.5459, 0.4475], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5972, 0.5312, 0.5822, 0.4033, 0.34, 0.4753, 0.4832, 0.5622, 0.6709, 0.4719, 0.5446, 0.6293, 0.6253, 0.6676, 0.5987, 0.6433, 0.517, 0.5319, 0.5642, 0.5824, 0.5984, 0.3935, 0.5031, 0.5808, 0.6142, 0.4506, 0.5254, 0.6104, 0.5801, 0.5496, 0.6086, 0.5073, 0.4144, 0.6471, 0.5267], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5608, 0.6813, 0.4887, 0.4403, 0.6209, 0.5955, 0.5595, 0.553, 0.4271, 0.5205, 0.6232, 0.5502, 0.5397, 0.632, 0.6445, 0.5504, 0.537, 0.6392, 0.5279, 0.612, 0.5404, 0.5493, 0.5563, 0.5629, 0.462, 0.5455, 0.5072, 0.863, 0.583, 0.5483, 0.8482, 0.423, 0.578, 0.5676], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5877, 0.3981, 0.4308, 0.3942, 0.4797, 0.493, 0.6869, 0.5562, 0.655, 0.5366, 0.6385, 0.6315, 0.5808, 0.541, 0.81, 0.5025, 0.5723, 0.6025, 0.5183, 0.4235, 0.4464, 0.501, 0.5468, 0.5617, 0.5756, 0.6159, 0.6013, 0.7616, 0.5398, 0.5053, 0.3204, 0.6157, 0.6491], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4343, 0.5278, 0.462, 0.6397, 0.6238, 0.5879, 0.4452, 0.5639, 0.7013, 0.4872, 0.5807, 0.557, 0.5832, 0.5958, 0.5435, 0.6599, 0.5626, 0.573, 0.4649, 0.5184, 0.5788, 0.6494, 0.4242, 0.5999, 0.5689, 0.723, 0.6066, 0.6919, 0.6746, 0.3817, 0.5607, 0.5499], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5404, 0.5277, 0.4994, 0.3617, 0.3947, 0.2562, 0.4111, 0.3768, 0.3592, 0.3877, 0.3764, 0.4319, 0.4207, 0.4404, 0.4149, 0.4642, 0.4303, 0.3892, 0.4934, 0.3919, 0.3486, 0.356, 0.3424, 0.4713, 0.4224, 0.3911, 0.394, 0.3574, 0.4434, 0.3743, 0.4241], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.3808, 0.5347, 0.4369, 0.4229, 0.2778, 0.4722, 0.4285, 0.2731, 0.3657, 0.3439, 0.401, 0.5026, 0.4854, 0.4952, 0.4463, 0.3985, 0.339, 0.4048, 0.3713, 0.3826, 0.4061, 0.354, 0.4346, 0.4499, 0.4568, 0.4311, 0.4073, 0.28, 0.3823, 0.4523], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5169, 0.4073, 0.4422, 0.2833, 0.3887, 0.4446, 0.4385, 0.3596, 0.4543, 0.629, 0.3593, 0.5934, 0.4739, 0.4723, 0.484, 0.517, 0.5447, 0.3718, 0.3414, 0.3404, 0.2913, 0.4493, 0.4857, 0.4059, 0.3568, 0.4585, 0.4647, 0.4904, 0.4848], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5676, 0.4298, 0.3651, 0.4917, 0.6101, 0.4808, 0.4976, 0.4298, 0.616, 0.494, 0.6171, 0.6744, 0.4554, 0.6154, 0.4455, 0.6564, 0.5193, 0.4862, 0.2982, 0.3799, 0.4489, 0.6069, 0.4953, 0.431, 0.5435, 0.3992, 0.4598, 0.4955], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4822, 0.4407, 0.5528, 0.6471, 0.4461, 0.5804, 0.4562, 0.5864, 0.475, 0.5101, 0.6103, 0.5439, 0.5046, 0.3944, 0.5114, 0.6023, 0.6966, 0.385, 0.4647, 0.5035, 0.5409, 0.5283, 0.5564, 0.4783, 0.3117, 0.6296, 0.5031], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4432, 0.641, 0.5526, 0.6147, 0.6302, 0.5851, 0.719, 0.542, 0.5717, 0.5427, 0.6188, 0.4661, 0.5162, 0.5119, 0.5984, 0.5386, 0.5557, 0.5182, 0.7171, 0.537, 0.5957, 0.5683, 0.4485, 0.3559, 0.6382, 0.5446], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4541, 0.4715, 0.5942, 0.6061, 0.4821, 0.3845, 0.5177, 0.4006, 0.4378, 0.5378, 0.506, 0.2767, 0.3571, 0.5109, 0.6121, 0.3714, 0.5556, 0.4499, 0.4375, 0.5081, 0.4763, 0.339, 0.2755, 0.4396, 0.4513], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6457, 0.5499, 0.6566, 0.4761, 0.5065, 0.6171, 0.5749, 0.7005, 0.5851, 0.5177, 0.4095, 0.3917, 0.5137, 0.5785, 0.5273, 0.5338, 0.6885, 0.5239, 0.7483, 0.576, 0.4626, 0.3397, 0.6661, 0.6879], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6084, 0.699, 0.5535, 0.5465, 0.5667, 0.6045, 0.69, 0.548, 0.6433, 0.3891, 0.4814, 0.5866, 0.7876, 0.4595, 0.6485, 0.594, 0.5917, 0.654, 0.7626, 0.5294, 0.355, 0.5971, 0.595], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.7882, 0.5979, 0.6221, 0.5591, 0.5509, 0.5083, 0.5852, 0.5913, 0.4406, 0.4359, 0.5507, 0.6184, 0.453, 0.7086, 0.601, 0.5525, 0.6266, 0.548, 0.4665, 0.3666, 0.5542, 0.5621], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5737, 0.5698, 0.6055, 0.5822, 0.6102, 0.6908, 0.5796, 0.3982, 0.4208, 0.5966, 0.6844, 0.4967, 0.7241, 0.6473, 0.5282, 0.6456, 0.6428, 0.4546, 0.362, 0.6501, 0.6245], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5896, 0.5641, 0.4112, 0.5695, 0.5741, 0.5414, 0.4616, 0.4296, 0.5073, 0.5144, 0.7818, 0.6048, 0.5708, 0.6254, 0.5654, 0.5493, 0.5846, 0.3313, 0.6438, 0.572], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4608, 0.6733, 0.5633, 0.5867, 0.5393, 0.645, 0.6189, 0.5524, 0.5527, 0.4575, 0.4835, 0.6085, 0.6352, 0.4904, 0.4876, 0.5756, 0.4571, 0.6628, 0.4861], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4813, 0.6464, 0.5522, 0.6124, 0.3652, 0.4238, 0.472, 0.5436, 0.5682, 0.5578, 0.4995, 0.612, 0.7659, 0.5582, 0.5427, 0.3128, 0.5725, 0.7244], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5578, 0.5627, 0.4699, 0.5546, 0.4965, 0.4402, 0.4725, 0.3528, 0.5048, 0.5388, 0.4875, 0.4933, 0.5407, 0.4628, 0.4083, 0.585, 0.4845], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4814, 0.6497, 0.4558, 0.4489, 0.503, 0.5683, 0.4968, 0.4907, 0.5215, 0.6602, 0.7268, 0.5328, 0.6074, 0.3531, 0.6096, 0.7096], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5059, 0.4326, 0.4732, 0.5556, 0.57, 0.5008, 0.5873, 0.6608, 0.485, 0.5664, 0.559, 0.4623, 0.377, 0.6599, 0.5363], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5118, 0.544, 0.5785, 0.5881, 0.4035, 0.5446, 0.4808, 0.5917, 0.5863, 0.5597, 0.518, 0.4837, 0.5154, 0.598], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5234, 0.3922, 0.3783, 0.3441, 0.37, 0.4185, 0.4818, 0.4233, 0.3874, 0.4578, 0.5316, 0.3837, 0.4097], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5461, 0.4588, 0.2912, 0.397, 0.436, 0.4987, 0.4327, 0.4502, 0.4338, 0.5666, 0.4507, 0.4195], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6868, 0.3773, 0.5392, 0.5479, 0.5749, 0.5011, 0.5225, 0.5028, 0.3199, 0.5607, 0.4458], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4239, 0.6759, 0.5906, 0.5858, 0.5959, 0.7133, 0.4717, 0.3412, 0.6078, 0.5173], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.495, 0.5604, 0.4786, 0.5677, 0.485, 0.4338, 0.2601, 0.5783, 0.5466], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5558, 0.5792, 0.5516, 0.6569, 0.517, 0.3137, 0.5458, 0.4912], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4744, 0.6189, 0.6643, 0.394, 0.4086, 0.6602, 0.5333], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6135, 0.5124, 0.8955, 0.3999, 0.5945, 0.5259], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.6024, 0.5318, 0.3695, 0.6122, 0.7838], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.4474, 0.3709, 0.569, 0.4935], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.3651, 0.5002, 0.5007], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.361, 0.3915], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 0.5757], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0]]
txt_CI = []

# Calculate CI
for filename in data_dict:
    CI = 0
    for i in range(0, 60):
        for j in range(i, 60):
            CI += data_dict[filename][i] * data_dict[filename][j] * similarity[i][j]
    data_dict[filename] = CI

# print(data_dict)
print(data_dict['713671-01-000170-0.txt'])

for index1 in range(len(filename_column_list)):
    for index2 in range(len(filename_column_r_list)):
        parts = str(filename_column_r_list[index2]).split('-')
        if parts[0] == str(filename_column_list[index1]) and int(parts[1]) + 2000 == year_column_list[index1]:
            txt_CI.append(data_dict[filename_column_r_list[index2]])

print(len(txt_CI))
print(txt_CI)

# Create a DataFrame for CI values and save to Excel
txt_CI_df = pd.DataFrame(txt_CI, columns=['CI'], index=filename_column_list)
output_excel_path = r"D:\excel\finance\txt_CI_60.xlsx"
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    txt_CI_df.to_excel(writer)
print('done')