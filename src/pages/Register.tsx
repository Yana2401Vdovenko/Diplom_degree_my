import "../styles/Register.css";
import React, { useState, useRef, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import {
  FaTimes,
  FaEnvelope,
  FaLock,
  FaPhone,
  FaEye,
  FaEyeSlash,
  FaImage,
  FaChalkboardTeacher,
  FaFlask,
  FaUserTie,
  FaArrowRight,
  FaArrowLeft,
  FaInfoCircle,
  FaCheckCircle,
  FaTrashAlt,
  FaSearch,
  FaChevronDown,
  FaUniversity,
  FaCalendarAlt,
  FaClock,
  FaChartBar,
  FaCogs,
  FaUser,
  FaBriefcase,
  FaShieldAlt,
  FaUserPlus
} from "react-icons/fa";
import { supabase } from "../utils/supabaseClient";

type Role = "teacher" | "lab" | "head";

function Register() {
  const navigate = useNavigate();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [currentStep, setCurrentStep] = useState<number>(1);
  const [selectedRole, setSelectedRole] = useState<Role>("head");
  const [showPass, setShowPass] = useState<boolean>(false);
  const [avatarPreview, setAvatarPreview] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [submitting, setSubmitting] = useState<boolean>(false);

  // Дані з БД
  type DbItem = { code: string; name: string; facultyCode?: string };
  const [departments, setDepartments] = useState<DbItem[]>([]);
  const [disciplines, setDisciplines] = useState<DbItem[]>([]);
  const [faculties, setFaculties] = useState<DbItem[]>([]);
  const [positions, setPositions] = useState<DbItem[]>([]);
  const [degrees, setDegrees] = useState<DbItem[]>([]);
  const [titles, setTitles] = useState<DbItem[]>([]);

  useEffect(() => {
    const loadRef = (table: string, codeCol: string, nameCol: string, setter: (items: DbItem[]) => void) => {
      void supabase.from(table).select(`${codeCol}, ${nameCol}`).order(nameCol).then(({ data, error }) => {
        if (!error) setter(((data ?? []) as unknown as Record<string, unknown>[]).map((r) => ({ code: String(r[codeCol] ?? "").trim(), name: String(r[nameCol] ?? "").trim() })));
      });
    };
    void supabase.from("Кафедра").select("Код_кафедри,Назва_кафедри,Код_факультету_нні").order("Назва_кафедри").then(({ data, error }) => {
      if (error) return;
      const rows = ((data ?? []) as unknown as Record<string, unknown>[]).map((row) => ({
        code: String(row["Код_кафедри"] ?? "").trim(),
        name: String(row["Назва_кафедри"] ?? "").trim(),
        facultyCode: String(row["Код_факультету_нні"] ?? "").trim(),
      })).filter((row) => row.code);
      setDepartments(rows);
      const first = rows[0];
      if (first) {
        setTeacherData((current) => ({ ...current, department: current.department || first.code }));
        setLabData((current) => ({ ...current, department: current.department || first.code }));
        setHeadData((current) => ({ ...current, department: current.department || first.code, faculty: current.faculty || first.facultyCode || "" }));
      }
    });
    loadRef("Дисципліни", "Код_дисципліни", "Назва_дисципліни", setDisciplines);
    loadRef("Факультет_ННІ", "Код_факультету_нні", "Назва_факультету_нні", setFaculties);
    loadRef("Посада", "Код_посади", "Назва_посади", setPositions);
    loadRef("Статус_викладача", "Код_статусу", "Назва_статусу", setDegrees);
    loadRef("Тип_викладача", "Код_типу", "Назва_типу", setTitles);
  }, []);

  // Крок 2: Основна інформація
  const [formData, setFormData] = useState({
    firstName: "",
    lastName: "",
    middleName: "",
    email: "",
    phone: "",
    password: "",
    confirmPassword: "",
    interfaceLang: "Українська",
    avatar: null as File | null
  });

  // Крок 3: Додаткові дані для Викладача
  const [teacherData, setTeacherData] = useState({
    degree: "",
    title: "",
    position: "",
    department: "",
    experience: "",
    rate: "1.0",
    maxLoad: "",
    office: "",
    workPhone: ""
  });

  // Крок 3: Додаткові дані для Лаборанта
  const [labData, setLabData] = useState({
    department: "",
    laboratory: "",
    workPhone: "",
    responsibilities: {
      equipment: false,
      network: false,
      multimedia: false,
      materials: false
    }
  });

  // Крок 3: Додаткові дані для Завідувача кафедри
  const [headData, setHeadData] = useState({
    department: "",
    faculty: "",
    degree: "",
    title: "",
    workPhone: "",
    office: ""
  });

  // Крок 4: Підтвердження
  const [agreements, setAgreements] = useState({
    terms: false,
    privacy: false,
    accuracy: false
  });

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
    if (errorMessage) setErrorMessage("");
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (file.size > 5 * 1024 * 1024) {
        setErrorMessage("Розмір фото не повинен перевищувати 5 МБ");
        return;
      }
      setFormData((prev) => ({ ...prev, avatar: file }));
      const reader = new FileReader();
      reader.onloadend = () => setAvatarPreview(reader.result as string);
      reader.readAsDataURL(file);
    }
  };

  const handleRemovePhoto = (e: React.MouseEvent) => {
    e.stopPropagation();
    setFormData((prev) => ({ ...prev, avatar: null }));
    setAvatarPreview(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const handleContinue = async (e: React.MouseEvent<HTMLButtonElement>) => {
    e.preventDefault();
    if (currentStep === 1) {
      setCurrentStep(2);
    } else if (currentStep === 2) {
      if (!formData.firstName.trim() || !formData.lastName.trim() || !formData.email.trim() || !formData.phone.trim()) {
        setErrorMessage("Ви не внесли всі необхідні дані");
        return;
      }
      if (!formData.password || formData.password.length < 6) {
        setErrorMessage("Пароль має містити щонайменше 6 символів");
        return;
      }
      if (formData.password !== formData.confirmPassword) {
        setErrorMessage("Паролі не збігаються");
        return;
      }
      setErrorMessage("");
      setCurrentStep(3);
    } else if (currentStep === 3) {
      setCurrentStep(4);
    } else if (currentStep === 4) {
      if (!agreements.terms || !agreements.privacy || !agreements.accuracy) {
        setErrorMessage("Необхідно погодитися з усіма умовами");
        return;
      }
      setErrorMessage("");
      setSubmitting(true);

      const fullName = `${formData.lastName} ${formData.firstName} ${formData.middleName}`.trim();

      const roleValue = selectedRole === "lab" ? "laborant" : selectedRole;
      const selectedDepartmentId = (selectedRole === "lab" ? labData.department : selectedRole === "head" ? headData.department : teacherData.department).trim();
      const selectedDepartment = departments.find((department) => department.code === selectedDepartmentId);
      if (!selectedDepartmentId) {
        setSubmitting(false);
        setErrorMessage("Оберіть кафедру для облікового запису.");
        return;
      }

      const { data, error } = await supabase.auth.signUp({
        email: formData.email.trim(),
        password: formData.password,
        options: {
          data: {
            full_name: fullName,
            role: selectedRole === "lab" ? "Лаборант" : selectedRole === "head" ? "Завідувач кафедри" : "Викладач",
            role_code: roleValue,
            phone: formData.phone,
            department: selectedDepartment?.name ?? "",
            department_code: selectedDepartmentId,
            Код_кафедри: selectedDepartmentId,
            faculty: selectedDepartment?.facultyCode ?? "",
            Код_факультету_нні: selectedDepartment?.facultyCode ?? "",
          },
        },
      });

      setSubmitting(false);

      if (error) {
        setErrorMessage(error.message);
        return;
      }

      if (data.user && !data.session) {
        setErrorMessage("Акаунт створено! Перевірте пошту для підтвердження реєстрації.");
        return;
      }

      if (data.session) {
        try {
          await supabase.functions.invoke("set-user-role", { body: { role: roleValue } });
        } catch {
          // Edge Function недоступна — роль зберігається в user_metadata
        }

        const teacherCode = `V${Date.now().toString(36).toUpperCase()}`.slice(0, 10);
        let teacherPayload: Record<string, unknown> = {
          "Код_викладача": teacherCode,
          "ПІБ": fullName,
          "Email_викладача": formData.email.trim(),
          "Телефон_викладача": formData.phone,
        };

        if (selectedRole === "teacher") {
          teacherPayload = {
            ...teacherPayload,
            "Науковий_ступінь": teacherData.degree,
            "Вчене_звання": teacherData.title,
          };
        } else if (selectedRole === "head") {
          teacherPayload = {
            ...teacherPayload,
            "Науковий_ступінь": headData.degree,
            "Вчене_звання": headData.title,
          };
        } else if (selectedRole === "lab") {
          teacherPayload = {
            ...teacherPayload,
            "Науковий_ступінь": "",
            "Вчене_звання": "",
          };
        }

        const { error: dbError } = await supabase.from("Викладач").upsert(teacherPayload, { onConflict: "Код_викладача" });
        if (dbError) {
          console.error("Помилка збереження профілю в БД:", dbError.message);
        }

        const displayRole = roleValue === "laborant" ? "Лаборант" : roleValue === "teacher" ? "Викладач" : "Завідувач кафедри";
        localStorage.setItem("userRole", displayRole);
        localStorage.setItem("userDisplayName", fullName);
        localStorage.setItem("userEmail", formData.email.trim());
        navigate("/dashboard");
      }
    }
  };

  const handleBack = () => {
    if (currentStep > 1) setCurrentStep(currentStep - 1);
  };

  return (
    <div className="modal-overlay">
      <div className="register-modal">
        <button className="close-modal" onClick={() => navigate("/")}>
          <FaTimes />
        </button>

        {/* Навігація ліворуч */}
        <aside className="steps-sidebar">
          <div className={`sidebar-step-item ${currentStep === 1 ? "active" : currentStep > 1 ? "completed" : ""}`}>
            <span className="step-circle-icon">{currentStep > 1 ? <FaCheckCircle className="step-success-icon" /> : "1"}</span>
            <span className="step-label-text">Вибір ролі</span>
          </div>
          <div className={`sidebar-step-item ${currentStep === 2 ? "active" : currentStep > 2 ? "completed" : ""}`}>
            <span className="step-circle-icon">{currentStep > 2 ? <FaCheckCircle className="step-success-icon" /> : "2"}</span>
            <span className="step-label-text">Основна інформація</span>
          </div>
          <div className={`sidebar-step-item ${currentStep === 3 ? "active" : currentStep > 3 ? "completed" : ""}`}>
            <span className="step-circle-icon">{currentStep > 3 ? <FaCheckCircle className="step-success-icon" /> : "3"}</span>
            <span className="step-label-text">Додаткові дані</span>
          </div>
          <div className={`sidebar-step-item ${currentStep === 4 ? "active" : ""}`}>
            <span className="step-circle-icon">4</span>
            <span className="step-label-text">Підтвердження</span>
          </div>
        </aside>

        {/* Права частина з вмістом */}
        <main className="register-content">
          <header className="content-header-title">
            <h2>Створення акаунта</h2>
            <p>Створіть обліковий запис для доступу до системи управління кафедрою</p>
          </header>

          {errorMessage && (
            <div className="error-banner">
              <FaInfoCircle />
              <span>{errorMessage}</span>
            </div>
          )}

          {/* КРОК 1: Вибір ролі */}
          {currentStep === 1 && (
            <section className="form-step-wrapper">
              <h3>Вибір ролі</h3>
              <p className="role-desc">Оберіть вашу роль у системі</p>
              <div className="roles-layout-row">
                <div className={`role-select-card ${selectedRole === 'teacher' ? 'selected' : ''}`} onClick={() => setSelectedRole('teacher')}>
                  <div className="role-icon-box teacher"><FaChalkboardTeacher /></div>
                  <div className="role-card-text">
                    <h4>Викладач {selectedRole === 'teacher' && <FaCheckCircle className="role-check-icon" />}</h4>
                    <p>Викладання дисциплін, ведення занять, оцінювання студентів та методична робота</p>
                  </div>
                </div>

                <div className={`role-select-card ${selectedRole === 'lab' ? 'selected' : ''}`} onClick={() => setSelectedRole('lab')}>
                  <div className="role-icon-box lab"><FaFlask /></div>
                  <div className="role-card-text">
                    <h4>Лаборант {selectedRole === 'lab' && <FaCheckCircle className="role-check-icon" />}</h4>
                    <p>Підтримка лабораторій, облік обладнання та допомога у навчальному процесі</p>
                  </div>
                </div>

                <div className={`role-select-card ${selectedRole === 'head' ? 'selected' : ''}`} onClick={() => setSelectedRole('head')}>
                  <div className="role-icon-box head"><FaUserTie /></div>
                  <div className="role-card-text">
                    <h4>Завідувач кафедри {selectedRole === 'head' && <FaCheckCircle className="role-check-icon" />}</h4>
                    <p>Управління кафедрою, формування навантаження та звітність</p>
                  </div>
                </div>
              </div>
            </section>
          )}

          {/* КРОК 2: Основна інформація */}
          {currentStep === 2 && (
            <section className="form-step-wrapper">
              <h3>Основна інформація</h3>
              <div className="form-fields-grid-three">
                <div className="form-input-group">
                  <label>Прізвище <span className="required-star">*</span></label>
                  <input type="text" name="lastName" value={formData.lastName} onChange={handleInputChange} placeholder="Введіть прізвище" />
                </div>
                <div className="form-input-group">
                  <label>Ім'я <span className="required-star">*</span></label>
                  <input type="text" name="firstName" value={formData.firstName} onChange={handleInputChange} placeholder="Введіть ім'я" />
                </div>
                <div className="form-input-group">
                  <label>По батькові</label>
                  <input type="text" name="middleName" value={formData.middleName} onChange={handleInputChange} placeholder="Введіть по батькові" />
                </div>
                <div className="form-input-group">
                  <label>Email <span className="required-star">*</span></label>
                  <div className="field-with-icon-wrapper">
                    <FaEnvelope className="field-inner-icon" />
                    <input type="email" name="email" value={formData.email} onChange={handleInputChange} placeholder="ivanenko@nubip.edu.ua" />
                  </div>
                </div>
                <div className="form-input-group">
                  <label>Номер телефону <span className="required-star">*</span></label>
                  <div className="field-with-icon-wrapper">
                    <FaPhone className="field-inner-icon" />
                    <input type="text" name="phone" value={formData.phone} onChange={handleInputChange} placeholder="+380 (67) 123-45-67" />
                  </div>
                   </div>
                    </div>

              <div className="form-fields-grid-three">
                <div className="form-input-group">
                  <label>Пароль <span className="required-star">*</span></label>
                  <div className="field-with-icon-wrapper">
                    <FaLock className="field-inner-icon" />
                    <input type={showPass ? "text" : "password"} name="password" value={formData.password} onChange={handleInputChange} placeholder="•••••••" />
                    <button onClick={() => setShowPass(!showPass)} type="button" className="eye-toggle-btn">
                      {showPass ? <FaEyeSlash /> : <FaEye />}
                    </button>
                  </div>
                </div>
                <div className="form-input-group">
                  <label>Підтвердження пароля <span className="required-star">*</span></label>
                  <div className="field-with-icon-wrapper">
                    <FaLock className="field-inner-icon" />
                    <input type={showPass ? "text" : "password"} name="confirmPassword" value={formData.confirmPassword} onChange={handleInputChange} placeholder="•••••••" />
                  </div>
                </div>
                <div className="form-input-group photo-upload-wrapper">
                  <label>Фото профілю</label>
                  <input type="file" ref={fileInputRef} onChange={handleFileChange} accept="image/*" style={{ display: "none" }} />
                  <div className="dashed-upload-zone" onClick={() => fileInputRef.current?.click()}>
                    {avatarPreview ? (
                      <div className="avatar-preview-box">
                        <img src={avatarPreview} alt="Preview" className="img-preview-circle" />
                        <button type="button" className="remove-photo-icon-btn" onClick={handleRemovePhoto}>
                          <FaTrashAlt />
                        </button>
                      </div>
                    ) : (
                      <>
                        <FaImage className="upload-icon" />
                        <span>Завантажити фото <small>(до 5 МБ)</small></span>
                      </>
                    )}
                  </div>
                </div>
              </div>
            </section>
          )}

          {/* КРОК 3: Додаткові дані */}
          {currentStep === 3 && (
            <section className="form-step-wrapper">
              {selectedRole === "teacher" && (
                <>
                  <h3>Додаткові дані для викладача</h3>
                  <div className="form-two-columns-layout">
                    <div className="form-vertical-column">
                        <div className="form-input-group custom-select-box">
                          <label>Науковий ступінь</label>
                          <select value={teacherData.degree} onChange={(e) => setTeacherData({...teacherData, degree: e.target.value})}>
                            <option value="">Оберіть ступінь</option>
                            {degrees.map((d) => <option key={d.code} value={d.name}>{d.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                        <div className="form-input-group custom-select-box">
                          <label>Вчене звання</label>
                          <select value={teacherData.title} onChange={(e) => setTeacherData({...teacherData, title: e.target.value})}>
                            <option value="">Оберіть звання</option>
                            {titles.map((t) => <option key={t.code} value={t.name}>{t.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                        <div className="form-input-group custom-select-box">
                          <label>Посада</label>
                          <select value={teacherData.position} onChange={(e) => setTeacherData({...teacherData, position: e.target.value})}>
                            <option value="">Оберіть посаду</option>
                            {positions.map((p) => <option key={p.code} value={p.name}>{p.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                        <div className="form-input-group custom-select-box">
                          <label>Кафедра <span className="required-star">*</span></label>
                          <select value={teacherData.department} onChange={(e) => setTeacherData({...teacherData, department: e.target.value})}>
                            <option value="">Оберіть кафедру</option>
                            {departments.map((d) => <option key={d.code} value={d.code}>{d.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>

                      <div className="form-grid-inner-two">
  <div className="form-input-group">
    <label>Стаж викладацької діяльності (років)</label>
    <input 
      type="text" 
      value={teacherData.experience} 
      onChange={(e) => setTeacherData({...teacherData, experience: e.target.value})} 
      placeholder="Введіть кількість років" 
    />
  </div>
  
  <div className="form-input-group rate-select-wrapper">
    <label>Ставка <span className="required-star">*</span></label>
    <div className="select-relative-container">
      <select 
        value={teacherData.rate} 
        onChange={(e) => setTeacherData({...teacherData, rate: e.target.value})}
      >
        <option value="1.0">1.0</option>
        <option value="0.75">0.75</option>
        <option value="0.5">0.5</option>
      </select>
      <FaChevronDown className="select-dropdown-arrow" />
    </div>
  </div>
</div>

                      <div className="form-input-group">
                        <label>Максимальне навчальне навантаження (год/семестр) <span className="required-star">*</span></label>
                        <input type="text" value={teacherData.maxLoad} onChange={(e) => setTeacherData({...teacherData, maxLoad: e.target.value})} placeholder="Введіть кількість годин" />
                      </div>
                    </div>

                    <div className="form-vertical-column">
                        <div className="form-input-group">
                          <label>Дисципліни, які викладає</label>
                          <div className="search-input-field-wrapper margin-b-10">
                            <FaSearch className="inner-search-icon" />
                            <input type="text" placeholder="Пошук дисциплін..." />
                          </div>
                          <div className="checkboxes-scrollable-container">
                            {disciplines.length > 0 ? disciplines.map((disc) => (
                              <label key={disc.code} className="checkbox-row-element">
                                <input type="checkbox" value={disc.code} />
                                <span className="checkbox-label-title">{disc.name}</span>
                              </label>
                            )) : (
                              <p className="department-empty">Дисципліни не знайдено</p>
                            )}
                          </div>
                        </div>

                      <div className="inner-contacts-card-wrapper">
                        <label className="inner-card-header-label">Контактна інформація</label>
                        <div className="form-grid-inner-two">
                          <div className="form-input-group">
                            <label>Номер кабінету</label>
                            <input type="text" value={teacherData.office} onChange={(e) => setTeacherData({...teacherData, office: e.target.value})} placeholder="Введіть номер кабінету" />
                          </div>
                          <div className="form-input-group">
                            <label>Робочий телефон</label>
                            <input type="text" value={teacherData.workPhone} onChange={(e) => setTeacherData({...teacherData, workPhone: e.target.value})} placeholder="Введіть робочий телефон" />
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              )}

              {selectedRole === "lab" && (
                <>
                  <h3>Додаткові дані для лаборанта</h3>
                  <div className="form-two-columns-layout">
                    <div className="form-vertical-column">
                        <div className="form-input-group custom-select-box">
                          <label>Кафедра <span className="required-star">*</span></label>
                          <select value={labData.department} onChange={(e) => setLabData({...labData, department: e.target.value})}>
                            <option value="">Оберіть кафедру</option>
                            {departments.map((d) => <option key={d.code} value={d.code}>{d.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                      <div className="form-input-group custom-select-box">
                        <label>Лабораторія <span className="required-star">*</span></label>
                        <select value={labData.laboratory} onChange={(e) => setLabData({...labData, laboratory: e.target.value})}>
                          <option value="Комп'ютерна лабораторія №2">Комп'ютерна лабораторія №2</option>
                        </select>
                        <FaChevronDown className="select-dropdown-arrow" />
                      </div>
                      <div className="form-input-group">
                        <label>Робочий телефон <span className="required-star">*</span></label>
                        <input type="text" value={labData.workPhone} onChange={(e) => setLabData({...labData, workPhone: e.target.value})} />
                      </div>
                    </div>

                    <div className="form-vertical-column">
                      <div className="form-input-group">
                        <label>Відповідальність за ресурси</label>
                        <div className="checkboxes-vertical-stack">
                          <label className="checkbox-row-element">
                            <input type="checkbox" checked={labData.responsibilities.equipment} onChange={(e) => setLabData({...labData, responsibilities: {...labData.responsibilities, equipment: e.target.checked}})} />
                            <span className="checkbox-label-title">Комп'ютерне обладнання</span>
                          </label>
                          <label className="checkbox-row-element">
                            <input type="checkbox" checked={labData.responsibilities.network} onChange={(e) => setLabData({...labData, responsibilities: {...labData.responsibilities, network: e.target.checked}})} />
                            <span className="checkbox-label-title">Мережеве обладнання</span>
                          </label>
                          <label className="checkbox-row-element">
                            <input type="checkbox" checked={labData.responsibilities.multimedia} onChange={(e) => setLabData({...labData, responsibilities: {...labData.responsibilities, multimedia: e.target.checked}})} />
                            <span className="checkbox-label-title">Мультимедійна техніка</span>
                          </label>
                          <label className="checkbox-row-element">
                            <input type="checkbox" checked={labData.responsibilities.materials} onChange={(e) => setLabData({...labData, responsibilities: {...labData.responsibilities, materials: e.target.checked}})} />
                            <span className="checkbox-label-title">Лабораторні матеріали</span>
                          </label>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              )}

              {selectedRole === "head" && (
                <>
                  <h3>Додаткові дані для завідувача кафедри</h3>
                  <div className="form-two-columns-layout">
                    <div className="form-vertical-column">
                        <div className="form-input-group custom-select-box">
                          <label>Кафедра <span className="required-star">*</span></label>
                          <select value={headData.department} onChange={(e) => { const department = departments.find((item) => item.code === e.target.value); setHeadData({ ...headData, department: e.target.value, faculty: department?.facultyCode ?? "" }); }}>
                            <option value="">Оберіть кафедру</option>
                            {departments.map((d) => <option key={d.code} value={d.code}>{d.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                        <div className="form-input-group custom-select-box">
                          <label>Факультет / ННІ <span className="required-star">*</span></label>
                          <select value={headData.faculty} onChange={(e) => setHeadData({...headData, faculty: e.target.value})}>
                            <option value="">Оберіть факультет</option>
                            {faculties.map((f) => <option key={f.code} value={f.code}>{f.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                        <div className="form-input-group custom-select-box">
                          <label>Науковий ступінь</label>
                          <select value={headData.degree} onChange={(e) => setHeadData({...headData, degree: e.target.value})}>
                            <option value="">Оберіть ступінь</option>
                            {degrees.map((d) => <option key={d.code} value={d.name}>{d.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                        <div className="form-input-group custom-select-box">
                          <label>Вчене звання</label>
                          <select value={headData.title} onChange={(e) => setHeadData({...headData, title: e.target.value})}>
                            <option value="">Оберіть звання</option>
                            {titles.map((t) => <option key={t.code} value={t.name}>{t.name}</option>)}
                          </select>
                          <FaChevronDown className="select-dropdown-arrow" />
                        </div>
                      <div className="form-input-group">
                        <label>Робочий телефон <span className="required-star">*</span></label>
                        <input type="text" value={headData.workPhone} onChange={(e) => setHeadData({...headData, workPhone: e.target.value})} />
                      </div>
                      <div className="form-input-group">
                        <label>Кабінет</label>
                        <input type="text" value={headData.office} onChange={(e) => setHeadData({...headData, office: e.target.value})} />
                      </div>
                    </div>

                    <div className="form-vertical-column">
                      <div className="form-input-group">
                        <label>Системні можливості</label>
                        <div className="system-widgets-list-stack">
                          <div className="system-feature-badge-card">
                            <div className="feature-icon-box"><FaUniversity /></div>
                            <div className="feature-info-content">
                              <h5>Управління кафедрою</h5>
                              <p>Повний доступ до управління структурою, співробітниками та діяльністю кафедри</p>
                            </div>
                          </div>
                          <div className="system-feature-badge-card">
                            <div className="feature-icon-box"><FaCalendarAlt /></div>
                            <div className="feature-info-content">
                              <h5>Формування навантаження</h5>
                              <p>Планування та розподіл навчального навантаження викладачів</p>
                            </div>
                          </div>
                          <div className="system-feature-badge-card">
                            <div className="feature-icon-box"><FaClock /></div>
                            <div className="feature-info-content">
                              <h5>Підтвердження розкладу</h5>
                              <p>Затвердження змін та публікація розкладу навчальних занять</p>
                            </div>
                          </div>
                          <div className="system-feature-badge-card">
                            <div className="feature-icon-box"><FaChartBar /></div>
                            <div className="feature-info-content">
                              <h5>Доступ до звітності</h5>
                              <p>Перегляд аналітики, звітів та статистики діяльності кафедри</p>
                            </div>
                          </div>
                          <div className="system-feature-badge-card">
                            <div className="feature-icon-box"><FaCogs /></div>
                            <div className="feature-info-content">
                              <h5>Керування налаштуваннями</h5>
                              <p>Управління даними викладачів, ролями та доступами в системі</p>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              )}
            </section>
          )}

          {/* КРОК 4: Підтвердження */}
          {currentStep === 4 && (
            <section className="form-step-wrapper confirmation-layout-wrapper">
              <h3>Перевірте введені дані</h3>
              <p className="confirmation-desc">Будь ласка, перевірте правильність введеної інформації перед створенням акаунта</p>
              
              <div className="confirmation-cards-row">
                <div className="preview-info-card-box">
                  <div className="card-top-header-title">
                    <FaUser className="card-header-icon" />
                    <h4>Основна інформація</h4>
                  </div>
                  <div className="preview-rows-list-stack">
                    <div className="info-data-row"><span className="data-key-label">ПІБ</span><span className="data-val-text">{`${formData.lastName} ${formData.firstName} ${formData.middleName}`}</span></div>
                    <div className="info-data-row"><span className="data-key-label">Email</span><span className="data-val-text">{formData.email}</span></div>
                    <div className="info-data-row"><span className="data-key-label">Телефон</span><span className="data-val-text">{formData.phone}</span></div>
                    <div className="info-data-row"><span className="data-key-label">Мова інтерфейсу</span><span className="data-val-text">{formData.interfaceLang}</span></div>
                  </div>
                </div>

                <div className="preview-info-card-box">
                  <div className="card-top-header-title">
                    <FaBriefcase className="card-header-icon" />
                    <h4>Обрана роль</h4>
                  </div>
                  <div className="role-selected-preview-widget">
                    {selectedRole === "teacher" && (
                      <div className="inner-role-badge">
                        <div className="role-icon-box teacher"><FaChalkboardTeacher /></div>
                        <div className="inner-role-info"><h5>Викладач</h5><p>Викладання дисциплін, ведення занять...</p></div>
                      </div>
                    )}
                    {selectedRole === "lab" && (
                      <div className="inner-role-badge">
                        <div className="role-icon-box lab"><FaFlask /></div>
                        <div className="inner-role-info"><h5>Лаборант</h5><p>Підтримка лабораторій, облік обладнання...</p></div>
                      </div>
                    )}
                    {selectedRole === "head" && (
                      <div className="inner-role-badge">
                        <div className="role-icon-box head"><FaUserTie /></div>
                        <div className="inner-role-info"><h5>Завідувач кафедри</h5><p>Управління кафедрою, формування навантаження та звітність</p></div>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="preview-info-card-box full-width-card">
                <div className="card-top-header-title">
                  <FaUniversity className="card-header-icon" />
                  <h4>
                    {selectedRole === "teacher" && "Дані викладача"}
                    {selectedRole === "lab" && "Дані лаборанта"}
                    {selectedRole === "head" && "Дані завідувача кафедри"}
                  </h4>
                </div>
                
                <div className="preview-data-grid-two-cols">
                  {selectedRole === "head" && (
                    <>
                      <div className="info-data-row"><span className="data-key-label">Кафедра</span><span className="data-val-text">{departments.find((item) => item.code === headData.department)?.name || "—"}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Факультет</span><span className="data-val-text">{faculties.find((item) => item.code === headData.faculty)?.name || "—"}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Науковий ступінь</span><span className="data-val-text">{headData.degree}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Вчене звання</span><span className="data-val-text">{headData.title}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Робочий телефон</span><span className="data-val-text">{headData.workPhone}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Кабінет</span><span className="data-val-text">{headData.office}</span></div>
                    </>
                  )}
                  {selectedRole === "teacher" && (
                    <>
                      <div className="info-data-row"><span className="data-key-label">Кафедра</span><span className="data-val-text">{departments.find((item) => item.code === teacherData.department)?.name || "—"}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Посада</span><span className="data-val-text">{teacherData.position}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Ставка</span><span className="data-val-text">{teacherData.rate}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Кабінет</span><span className="data-val-text">{teacherData.office}</span></div>
                    </>
                  )}
                  {selectedRole === "lab" && (
                    <>
                      <div className="info-data-row"><span className="data-key-label">Кафедра</span><span className="data-val-text">{departments.find((item) => item.code === labData.department)?.name || "—"}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Лабораторія</span><span className="data-val-text">{labData.laboratory}</span></div>
                      <div className="info-data-row"><span className="data-key-label">Робочий телефон</span><span className="data-val-text">{labData.workPhone}</span></div>
                    </>
                  )}
                </div>
              </div>

              <div className="preview-info-card-box full-width-card legal-agreements-panel">
                <div className="card-top-header-title">
                  <FaShieldAlt className="card-header-icon" />
                  <h4>Підтвердження</h4>
                </div>
                <div className="checkboxes-vertical-stack font-size-13">
                  <label className="checkbox-row-element">
                    <input type="checkbox" checked={agreements.terms} onChange={(e) => setAgreements({...agreements, terms: e.target.checked})} />
                    <span className="checkbox-label-title">Погоджуюся з умовами використання системи</span>
                  </label>
                  <label className="checkbox-row-element">
                    <input type="checkbox" checked={agreements.privacy} onChange={(e) => setAgreements({...agreements, privacy: e.target.checked})} />
                    <span className="checkbox-label-title">Погоджуюся на обробку персональних даних</span>
                  </label>
                  <label className="checkbox-row-element">
                    <input type="checkbox" checked={agreements.accuracy} onChange={(e) => setAgreements({...agreements, accuracy: e.target.checked})} />
                    <span className="checkbox-label-title">Підтверджую достовірність наданої інформації</span>
                  </label>
                </div>
              </div>
            </section>
          )}

          {/* Навігація внизу */}
          <footer className="form-action-footer">
            <div className="left-footer-actions">
              {currentStep > 1 && (
                <button className="btn-action-back" type="button" onClick={handleBack}>
                  <FaArrowLeft /> Назад
                </button>
              )}
            </div>
      <div className="right-footer-actions">
        <button className="btn-action-cancel" type="button" onClick={() => navigate("/")}>Скасувати</button>
        <button className="btn-action-continue" type="button" onClick={handleContinue} disabled={submitting}>
          {currentStep === 4 ? (
            submitting ? "Створення…" : <>Створити акаунт <FaUserPlus /></>
          ) : (
            <>Продовжити <FaArrowRight /></>
          )}
        </button>
      </div>
          </footer>
        </main>
      </div>
    </div>
  );
}

export default Register;

