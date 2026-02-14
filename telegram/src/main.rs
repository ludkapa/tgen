use dotenvy::dotenv;
use std::{env, error::Error, net::SocketAddr};
use teloxide::{
    dispatching::dialogue::{GetChatId, InMemStorage},
    prelude::*,
    types::User,
    update_listeners::webhooks,
    utils::markdown::user_mention_or_link,
};

type HandlerResult = Result<(), Box<dyn Error + Sync + Send>>;
type UserDialogue = Dialogue<DState, InMemStorage<DState>>;

//Dialogue State
#[derive(Clone, Default)]
enum DState {
    #[default]
    Start,
    MainMenu,
    Salary,
    YearSelect {
        salary: String,
    },
}

#[tokio::main]
async fn main() {
    // Load logs
    pretty_env_logger::init();
    log::info!("Запуск бота...");
    // Load envs
    log::info!("Загрузка env...");
    dotenv().ok();
    let token = env::var("BOT_TOKEN").unwrap();
    let port = env::var("PORT").unwrap();
    let url = env::var("WEBHOOK_URL").unwrap();
    run_bot(token, port, url).await;
}

async fn run_bot(token: String, port: String, webhook_url: String) {
    // Init bot
    let bot = Bot::new(token);
    // Init ipv4 addr
    let addr = SocketAddr::from(([0, 0, 0, 0], port.parse::<u16>().unwrap()));
    // Setup webhook listener
    let listener = webhooks::axum(
        bot.clone(),
        webhooks::Options::new(addr, webhook_url.parse().unwrap()),
    )
    .await
    .expect("Не удалось поднять Webhook!");
    // Dialogue update logic
    let router = Update::filter_message()
        .enter_dialogue::<Message, InMemStorage<DState>, DState>()
        .branch(dptree::case![DState::Start].endpoint(start))
        .branch(dptree::case![DState::MainMenu].endpoint(main_menu))
        .branch(dptree::case![DState::Salary].endpoint(salary))
        .branch(dptree::case![DState::YearSelect { salary }].endpoint(year_select));
    // Dispatcher
    Dispatcher::builder(bot, router)
        .dependencies(dptree::deps![InMemStorage::<DState>::new()])
        .enable_ctrlc_handler()
        .build()
        .dispatch_with_listener(
            listener,
            LoggingErrorHandler::with_custom_text("Ошибка при обновлении!"),
        )
        .await;
    todo!()
}

async fn start(bot: Bot, dialogue: UserDialogue, msg: Message) -> HandlerResult {
    let user = msg.from;
    let user_name: String = match user {
        Some(user) => match user.username {
            Some(username) => username,
            None => user.id.0.to_string(),
        },
        None => "пользователь".to_string(),
    };
    bot.send_message(msg.chat.id, format!("Привет {}, этот бот генерирует табель для подсчёта переработок.\nВыбери что ты хочешь сделать ниже:", user_name)).await?;
    dialogue.update(DState::MainMenu).await?;
    Ok(())
}

async fn main_menu(bot: Bot, dialogue: UserDialogue, msg: Message) -> HandlerResult {
    todo!()
}
async fn salary(bot: Bot, dialogue: UserDialogue, msg: Message) -> HandlerResult {
    todo!()
}
async fn year_select(
    bot: Bot,
    dialogue: UserDialogue,
    salary: String,
    msg: Message,
) -> HandlerResult {
    todo!()
}
